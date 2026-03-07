Partial Class TC_05_001_chg
    Inherits AuthBasePage

#Region "DIM CONST 宣告"

    '70:區域產業據點職業訓練計畫(在職)
    Dim flag_TPlanID70_1 As Boolean = False ' (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)

    '技檢訓練時數 '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    Const cst_EHour_t1 As String = "技檢訓練時數,目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時可儲存，若不符合上述條件，該資料不會存入資料庫。"
    '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_Use_TMID As String="672"
    Const cst_DG3_EHour_技檢訓練時數_iCOL As Integer = 5
    Const cst_DG4_EHour_技檢訓練時數_iCOL As Integer = 5

    Dim flag_PackageType_NOUSE As Boolean = True 'true:(未使用)未選擇 包班種類(移除) 'false:(使用)選擇 包班種類(保留)
    '產投使用／遠距教學 暫不啟用
    Dim flag_StopDISTANCE As Boolean = True
    'Dim ChgItemName As String() '將變更項目名稱定義到陣列之中

    Dim MyValue As String = ""
    Const cst_labTitle_申請狀態 As String = "申請狀態"
    Const cst_labTitle_計畫狀態 As String = "計畫狀態"

    Const cst_OldData9_1_開辦中 As String = "開辦中"
    Const cst_OldData9_1_停辦中 As String = "停辦中"

    Const cst_SearchMode_申請 As String = "申請"
    Const cst_SearchMode_變更結果 As String = "變更結果"

    Const cst_CheckMode_審核中 As String = "審核中"
    Const cst_CheckMode_審核通過 As String = "審核通過"
    Const cst_CheckMode_審核不通過 As String = "審核不通過"

    Const cst_errmsg_alt1 As String = "【非學分班】，訓練起迄日期區間，不得超過4個月"
    'Const cst_errmsg_alt1b As String="訓練起迄日期區間，不得超過12個月"
    Const cst_errmsg_alt1c As String = "【非學分班】，訓練起迄日期區間，迄日不得超過次年04/30"
    Const cst_errmsg_alt2 As String = "訓練起迄日期有誤!"

    Const cst_errmsg_alt3 As String = "新的訓練日期不能與舊日期的相同!"
    Const cst_errmsg_alt4 As String = "請輸入變更內容的起迄日期!"
    Const cst_errmsg_alt4b As String = "請確認輸入變更內容的起迄日期格式!"
    Const cst_errmsg_alt5 As String = "當日已有申請資料，不可再次申請相同資料!"

    '1、申請階段於「上半年」之課程，【開訓日期】不得 > 6/30 ，【結訓日期】必須 <= 8/31 。
    '2、申請階段於「下半年」之課程，【開訓日期】不得 > 當年度 12/31 ，【結訓日期】必須 <= 2月底 。
    '3、當年度申請階段於「政策性產業」之課程，結訓日不得超過翌年 4/30。(申請階段屬政策性產業，開訓日期可在翌年)
    '1、申請階段於「上半年」之課程，開訓日期不得超過6月30日(結訓日最遲在8月31日止)。
    '2、申請階段於「下半年」之課程，開訓日期不得超過當年度12月31日(結訓日最遲在翌年2月底止)。
    '3、當年度申請階段於「政策性產業」之課程，結訓日不得超過翌年4月30日。(申請階段屬政策性產業，開訓日期可在翌年)
    Const cst_errmsg_alt61 As String = "申請階段於「上半年」之課程，開訓日期不得超過6月30日(結訓日最遲在8月31日止)"
    Const cst_errmsg_alt62 As String = "申請階段於「下半年」之課程，開訓日期不得超過當年度12月31日(結訓日最遲在翌年2月底止)"
    Const cst_errmsg_alt63 As String = "申請階段於「政策性產業」之課程，結訓日不得超過翌年4月30日。"
    Const cst_errmsg_alt3b As String = "申請階段於「進階政策性產業」之課程，結訓日期(訓練迄日),必須在當年度!"

    Const cst_i_ReviseCont_c_max_length As Integer = 250
    Const cst_i_Times_c_max_length As Integer = 200 'by AMU 20220310
    Const cst_i_Times_c_min_length As Integer = 5

    Const cst_now As String = "now" 'sType@now新版課表(產投)
    Const cst_old1 As String = "old1" 'sType@old1舊1課表(產投)
    Const cst_btnDel1Cmd As String = "btnDel1"
    Const cst_NNN As String = "NNN"

    Dim rPlanID As Integer '計畫PK
    Dim rComIDNO As String '計畫PK
    Dim rSeqNo As Integer '計畫PK
    Dim rSCDate As String '變更PK
    Dim iSubSeqNO As Integer = 0 '變更PK(INT) '此變數有重複宣告可能
    Dim rAltDataID As String '變更 應為數字 chgState.Value =val(rAltDataID )
    Dim rPARTREDUC1 As String = "" 'Y OJT-21080202：產投 -班級變更申請：新增修改功能 PARTREDUC

    Dim i_gSeqno As Integer = 0 '共用序號使用
    Dim sWOScript1 As String = "" '共用JS OPEN語法

    'Const cst_inline1 As String="inline"
    Const cst_inline1 As String = ""
    Dim dtlist11 As DataTable = Nothing '師資(產投)
    Dim dtlist20 As DataTable = Nothing '助教(產投)

    Const vs_OCID As String = "OCID"
    Const vs_UpdateTrainDesc As String = "UpdateTrainDesc" 'ViewState(vs_UpdateTrainDesc)
    'Const vs_Do_CreateTrainDesc As String="Do_CreateTrainDesc" 'ViewState(vs_Do_CreateTrainDesc)
    Const vs_dtTaddress As String = "dtTaddress" 'ViewState(vs_dtTaddress)
    Const vs_IsLoaded As String = "IsLoaded" 'ViewState(vs_IsLoaded)
    Const vs_UpdateItemIndex As String = "UpdateItemIndex" 'ViewState(vs_UpdateItemIndex)
    Const vs_SubSeqNO As String = "_SubSeqNO" 'ViewState("_SubSeqNO

    Const vs_TEMP11_TrainDescDT As String = "TEMP11_TrainDescDT" 'ViewState(vs_TEMP11_TrainDescDT) 
    Const vs_TEMP20_TrainDescDT As String = "TEMP20_TrainDescDT" 'ViewState(vs_TEMP20_TrainDescDT)
    '產生暫存新的訓練日期
    Const cst_ss_TEMP1_TrainDescDT As String = "TEMP1_TrainDescDT"
    Const cst_PointYN_非學分班 As String = "非學分班"
    Const cst_PointYN_學分班 As String = "學分班"
    'Const vs_ChgItemUpdateItem As String="ChgItemUpdateItem" 'ViewState(vs_ChgItemUpdateItem)
    'Dim SCDate As String '變更PK (新增)
    'Dim SubSeqNO As Integer '變更PK  (新增)

    'AltDataID 變更項目
    '直接更動 aspx

    'AltDataID 變更項目 '於1:「開結訓日」、15:「上課時間」、14:「上課地點」、9:「停辦」等變更項目，新增「其他應備文件」欄位，放在公文項目前面，    
    '產投選項，職前選項自行增減
    'ChgItem=TIMS.TPlanID28ChgItemName
    'ChgItem [直接改介面]
    Const Cst_i訓練期間 As Integer = 1 '開、結訓日期(產投)
    Const Cst_i訓練時段 As Integer = 2
    Const Cst_i訓練地點 As Integer = 3
    Const Cst_i課程編配 As Integer = 4
    Const Cst_i訓練師資 As Integer = 5
    Const Cst_i班別名稱 As Integer = 6
    Const Cst_i期別 As Integer = 7
    Const Cst_i上課地址 As Integer = 8
    Const Cst_i停辦 As Integer = 9 '停辦
    Const Cst_i上課時段 As Integer = 10
    Const Cst_i師資 As Integer = 11 '師資(產投)
    Const Cst_i助教 As Integer = 20 '助教(產投) '20120213 BY AMU (產投用助教)

    Const Cst_i核定人數 As Integer = 12  'Cst_招生人數  as Integer=12
    Const Cst_i增班 As Integer = 13
    Const Cst_i科場地 As Integer = 14 '上課地點(產投)／學(術)科場地
    Const Cst_i上課時間 As Integer = 15 '上課時間(產投)
    Const Cst_i其他 As Integer = 16
    Const Cst_i報名日期 As Integer = 17  '20080825 andy  add 報名日期
    Const Cst_i課程表 As Integer = 18  '20080626 andy add 課程表(產投)

    Const Cst_i包班種類 As Integer = 19  '20111208 BY AMU 
    Const Cst_i訓練費用 As Integer = 21  '20170908 (職前)
    Const Cst_i遠距教學 As Integer = 22  '2021/06/09'增修需求 OJT-21060201 產投 - 班級變更申請/審核：新增遠距教學變更 + 網站-顯示遠距教學資訊 DISTANCE learning /distance teaching
    Const Cst_iMaxChgItem As Integer = 22

    Const Cst_sql As String = "SQL"
    Const Cst_msg1 As String = "請記得於本變更申請審核通過後，修正課程表。\n"
    'Const Cst_msg2 As String="原無師資2資料，不提供師資2資料變更!"
    Const Cst_msg2 As String = "原無 助教1 資料，不提供 助教1 資料變更!"
    Const Cst_msg3 As String = "原無 助教2 資料，不提供 助教2 資料變更!"

    Dim rActCheck As String = "" 'Request("check") 'Cst_cPlan:申請; Cst_cRevise:變更結果
    Const Cst_cPlan As String = "PLAN_PLANINFO" '申請
    Const Cst_cRevise As String = "PLAN_REVISE" '變更結果

    Dim vsShowmsg4 As String = "" '動作記錄功能，避免重複sql動作
    Dim flagDebugTest As Boolean = True '(啟動錯誤當機功能)

    Dim sPCS1 As String = "" '本班組合後的PCS PxCxS,
    Dim s_SPEC_PCSs1 As String = "" 'spec_PCSs1=TIMS.Utl_GetConfigSet("spec_PCSs1") '某些班級可使用特殊規則1。

    'Dim au As New cAUTH
    Dim ff33 As String = ""
    Dim strTMP1 As String = ""
    Dim dt_KEY_COSTITEM As DataTable = Nothing

    Const cst_TC05001CHG_COSTITEM_GUID As String = "TC05001CHG_COSTITEM_GUID"
    Dim gobjconn As SqlConnection

#End Region

#Region "REM"

    'insert
    'update
    'delete

    'PLAN_TRAINDESC 申請表
    'PLAN_TRAINDESC_REVISE '(某次申請變更)
    'PLAN_TRAINDESC_REVISEITEM '(申請變更的細項)
    'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
    'PLAN_TRAINDESC_RO (存申請舊表)
    'REVISE_TEACHER
    '----X-PLAN_TRAINDESC_REVISEOLD (已停用，不使用該表單)

    'DataGrid3_ItemDataBound

    '產投
    'SELECT * FROM PLAN_ONCLASS WHERE PlanID='1844' 
    'SELECT * FROM REVISE_ONCLASS WHERE PlanID='1844'
    'select * from PLAN_PLANINFO where PlanID=1801
    'select * from CLASS_CLASSINFO where PlanID=1801
    'select * from PLAN_REVISE where PlanID=1801
    'select * from PLAN_TRAINDESC_REVISE where PlanID=1801
    'select * from PLAN_TRAINDESC_REVISEITEM m where exists ( select 'x' from PLAN_TRAINDESC_REVISE c where m.PTDRID=c.PTDRID and c.PlanID=1801 )
    'select * from PLAN_TRAINDESC where PlanID=1801
    'select b.* from Teach_TeacherInfo b join PLAN_TRAINDESC a on a.techid=b.techid where a.PlanID=1801
    'select sc.* from CLASS_SCHEDULE sc 
    'join CLASS_CLASSINFO cc On sc.ocid =cc.ocid
    'join PLAN_PLANINFO pp on pp.planid =cc.planid and pp.comidno =cc.comidno and pp.seqno =cc.seqno
    'where 1=1 AND pp.PlanID=1801
    'select count(*) cnt  from Course_CourseInfo m where 1=1
    'and exists ( select 'x' from PLAN_PLANINFO pp join org_orginfo oo on oo.comidno =pp.comidno where 1=1  and oo.orgid =m.orgid and pp.PlanID=1801 )
    'select a.*  'from Teach_TeacherInfo a 'where exists (
    '	select 'x' from PLAN_PLANINFO pp  where 1=1 and pp.rid= a.RID and pp.PlanID=1728 --and pp.comidno ='45668746' and pp.seqno='6'    ')
    'select * FROM PLAN_TRAINPLACE  'where comidno ='45668746' 

#End Region

#Region "Public Shared"

    'PLAN_REVISE Show_PlanRevise TC_05_001_chg.Get_PlanReviseDataRow
    Public Shared Function Get_PlanReviseDataRow(ByRef htSS As Hashtable, ByRef oConn As SqlConnection) As DataRow
        'Dim htSS As New Hashtable 'htSS Hashtable() 'htSS.Add("strSetId", strSetId)
        'ByRef htSS As Hashtable 'Dim htSS As New Hashtable 'htSS Hashtable() 'htSS.Add("strSetId", strSetId)
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") 'Request("PlanID")
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") 'Request("cid")
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") 'Request("no")
        Dim rCDate As String = TIMS.GetMyValue2(htSS, "rCDate") 'Request("CDate")
        Dim rSubNo As String = TIMS.GetMyValue2(htSS, "rSubNo") 'Request("SubNo")

        Dim rst As DataRow = Nothing
        Call TIMS.OpenDbConn(oConn)
        Dim sql As String = ""
        sql &= " SELECT * FROM PLAN_REVISE" & vbCrLf
        sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO" & vbCrLf
        sql &= " AND CDate=@CDate AND SubSeqNo=@SubSeqNo" & vbCrLf
        'sql &= "AND CDate=convert(datetime, @CDate, 111)" & vbCrLf
        Dim odt As New DataTable

        'parms.Clear()
        Dim parms As Hashtable = New Hashtable From {
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNO", rSeqNo},
            {"CDate", rCDate},
            {"SubSeqNo", rSubNo}
        }
        odt = DbAccess.GetDataTable(sql, oConn, parms)
        If odt.Rows.Count > 0 Then rst = odt.Rows(0) '若查無資料為nothing
        Return rst
    End Function

    '檢核 排課作業 是否完成!!
    Public Shared Function CheckClassSchedule(ByVal sOCID As String, ByRef sErrmsg As String, ByVal oConn As SqlConnection) As Boolean
        Dim Rst As Boolean = False
        sErrmsg = "本班目前尚未排課，請先確認是否已於課程管理完成排課作業！-236A"
        If Not IsNumeric(sOCID) Then Return Rst
        Dim sql As String = " SELECT 'X' FROM CLASS_SCHEDULE WHERE OCID=@OCID "
        Dim parms As Hashtable = New Hashtable From {{"OCID", sOCID}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
        If dt.Rows.Count > 0 Then
            Rst = True
            sErrmsg = ""
        End If
        Return Rst
    End Function

    ''' <summary>整理異常資料並刪除(停止刪除動作)</summary>
    ''' <param name="iType"></param>
    ''' <param name="tmpID"></param>
    ''' <param name="tmpCOMIDNO"></param>
    ''' <param name="tmpSNO"></param>
    ''' <param name="tmpDate"></param>
    ''' <param name="tmpSubSNO"></param>
    ''' <param name="sAltDataID"></param>
    ''' <param name="oConn"></param>
    Public Shared Function Get_PTDRIDxDelErr(ByRef sm As SessionModel, ByVal iType As Integer, ByVal tmpID As Integer, ByVal tmpCOMIDNO As String, ByVal tmpSNO As Integer, ByVal tmpDate As String, ByVal tmpSubSNO As Integer, ByVal sAltDataID As String, ByVal oConn As SqlConnection) As Boolean
        'iType :1 為申請 /2 為變更結果
        Call TIMS.OpenDbConn(oConn)
        Dim rst As Boolean = True 'Return rst'

        Dim sql_1 As String = ""
        sql_1 &= " SELECT p.PTDRID" & vbCrLf '/*PK*/ 'sql &= " ,p.PLANID ,p.COMIDNO ,p.SEQNO" & vbCrLf 'sql &= " ,p.CDATE ,p.SUBSEQNO" & vbCrLf
        sql_1 &= " FROM PLAN_TRAINDESC_REVISE p" & vbCrLf
        sql_1 &= " WHERE p.PlanID=@planid" & vbCrLf
        sql_1 &= " AND p.ComIDNO=@comidno" & vbCrLf
        sql_1 &= " AND p.SeqNO=@seqno" & vbCrLf
        sql_1 &= " AND p.CDate=@cdate" & vbCrLf
        sql_1 &= " AND p.SubSeqNO=@subseqno" & vbCrLf
        Dim dt_1 As New DataTable
        Dim sCmd_1 As New SqlCommand(sql_1, oConn)
        With sCmd_1
            .Parameters.Clear()
            .Parameters.Add("planid", SqlDbType.Int).Value = TIMS.GetValue1(tmpID)
            .Parameters.Add("comidno", SqlDbType.VarChar).Value = TIMS.GetValue1(tmpCOMIDNO)
            .Parameters.Add("seqno", SqlDbType.Int).Value = TIMS.GetValue1(tmpSNO)
            .Parameters.Add("cdate", SqlDbType.DateTime).Value = CDate(TIMS.Cdate2(tmpDate))
            .Parameters.Add("subseqno", SqlDbType.Int).Value = tmpSubSNO
            '.Parameters.Add("AltDataID", SqlDbType.VarChar).Value=sAltDataID
            'dt_1=DbAccess.GetDataTable(sCmd_1.CommandText, oConn, sCmd_1.Parameters)
            dt_1.Load(.ExecuteReader())
        End With
        If dt_1.Rows.Count = 0 Then Return rst 'Exit Sub '無資料離開
        If dt_1.Rows.Count = 1 AndAlso iType = 2 Then
            '2:變更結果(查詢)，應該只能有1筆資料/或沒資料
            Return rst 'Exit Sub '有1筆資料離開
        End If
        '大於1筆資料-異常-課表沒有安排申請屬性
        If dt_1.Rows.Count > 1 Then
            rst = False '有資料為異常（退出） 
            Return rst '有資料為異常 
        End If

        Dim tmpPTDRIDs_IN As String = ""
        If dt_1.Rows.Count > 0 Then
            For Each dr As DataRow In dt_1.Rows
                tmpPTDRIDs_IN &= String.Concat(If(tmpPTDRIDs_IN <> "", ",", ""), dr("PTDRID"))
            Next
        End If
        If tmpPTDRIDs_IN = "" Then Return rst ' Exit Sub

        Dim sql_2 As String = ""
        sql_2 &= " SELECT x.PTDRIID ,x.PTDRID ,x.PTDID ,x.ALTDATAID ,x.ALTDATAITEM ,x.OLDDATA ,x.NEWDATA" & vbCrLf
        sql_2 &= " FROM PLAN_TRAINDESC_REVISEITEM x" & vbCrLf
        sql_2 &= String.Concat(" WHERE x.PTDRID IN (", tmpPTDRIDs_IN, ")", vbCrLf)
        Dim dt_2 As New DataTable
        Dim sCmd_2 As New SqlCommand(sql_2, oConn)
        With sCmd_2
            .Parameters.Clear()
            dt_2.Load(.ExecuteReader())
        End With
        If dt_2.Rows.Count = 0 Then Return rst 'Exit Sub '無資料離開

        '有資料為異常 'Common.MessageBox(MyPage, cst_errmsg_alt5)  Exit Sub
        If dt_2.Rows.Count > 0 Then
            rst = False '有資料為異常（退出） 
            Return rst '有資料為異常 
        End If

        '刪除 DELETE PLAN_TRAINDESC_REVISEITEM. DELETE PLAN_TRAINDESC_REVISE.
        Call DEL_PLAN_TRAINDESC_REVISE(sm, dt_2, sAltDataID, oConn)
    End Function

    ''' <summary>DELETE PLAN_TRAINDESC_REVISE // DELETE PLAN_TRAINDESC_REVISEITEM </summary>
    ''' <param name="sm"></param>
    ''' <param name="dt_2"></param>
    ''' <param name="sAltDataID"></param>
    ''' <param name="oConn"></param>
    Public Shared Sub DEL_PLAN_TRAINDESC_REVISE(ByRef sm As SessionModel, ByRef dt_2 As DataTable, ByRef sAltDataID As String, ByRef oConn As SqlConnection)
        If dt_2.Rows.Count = 0 Then Return
        For Each dr As DataRow In dt_2.Rows
            '可能異常資料(申請不可有當日的資料)
            If sAltDataID <> TIMS.CINT1(dr("ALTDATAID")) Then
                '刪除 DELETE PLAN_TRAINDESC_REVISEITEM. DELETE PLAN_TRAINDESC_REVISE.
                Call TIMS.DEL_PLAN_TRAINDESC_REVISEITEM(sm, TIMS.CINT1(dr("PTDRID")), oConn)
            End If
        Next
    End Sub

    ''' <summary> 報名日期(最後可報名時間) </summary>
    ''' <param name="MyValueSS"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Public Shared Function Get_SignUpEDateVal(ByVal MyValueSS As String, ByRef oConn As SqlConnection) As String
        Dim rst As String = ""
        Dim rPlanID As String = TIMS.GetMyValue(MyValueSS, "rPlanID")
        Dim rComIDNO As String = TIMS.GetMyValue(MyValueSS, "rComIDNO")
        Dim rSeqNo As String = TIMS.GetMyValue(MyValueSS, "rSeqNo")
        If rPlanID = "" Then Return rst
        If rComIDNO = "" Then Return rst
        If rSeqNo = "" Then Return rst

        '20081107 andy  add 報名日期
        Dim sql As String = ""
        sql &= " SELECT convert(varchar, STDate-1, 111) SignUpEDate "
        sql &= " FROM CLASS_CLASSINFO "
        sql &= " WHERE 0=0 "
        sql &= " AND PlanID=@PlanID " '" & rPlanID & "'"
        sql &= " AND ComIDNO=@ComIDNO " '" & rComIDNO & "'"
        sql &= " AND SeqNo=@SeqNo " '" & rSeqNo & "'"
        Dim sCmd As New SqlCommand(sql, oConn)
        Call TIMS.OpenDbConn(oConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.VarChar).Value = rPlanID
            .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = rComIDNO
            .Parameters.Add("SeqNo", SqlDbType.VarChar).Value = rSeqNo
            rst = Convert.ToString(.ExecuteScalar())
        End With
        Return rst
    End Function

#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(gobjconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        gobjconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(gobjconn)

        s_SPEC_PCSs1 = TIMS.Utl_GetConfigSet("spec_PCSs1") '某些班級可使用特殊規則。

        '每次重載執行
        Call Utl_EveryCreate1()

        '下拉選項，首頁資料顯示
        If Not IsPostBack Then
            cCreate1()
        End If

        '每次執行-隱藏所有資料列
        Call HidAllTr()
        '每次執行-顯示必要資料列 (有AUTOPOSTBACK造成)
        Call DisplayTR()

        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(gobjconn)) '每次執行 RegisterStartupScript
        'CTName.Attributes("onblur")="getzipname(this.value,'CTName','NewData8_1');"
    End Sub

    ''' <summary> 啟動儲存鈕機制 </summary>
    Sub Utl_CanSave_ButSub28()
        But_Sub28.Style("display") = cst_inline1 '(確定)按鈕
        But_Save28.Style("display") = "none"
        '課程表 針對產業人才投資計畫
        Select Case ViewState(vs_UpdateTrainDesc)
            Case "" '正式儲存=''  尚未按下確定鈕
            Case "Y"   '正式儲存='Y' ；當按下儲存鈕-->產生 1.計畫變更檔 2.課程表變更檔
                But_Sub28.Style("display") = "none"
                But_Save28.Style("display") = cst_inline1
                'But_UPLOAD28.Style("display")=cst_inline1 '準備進入上傳
            Case "N"   '正式儲存='N' 已按下確定鈕，(確定鈕--不可見，儲存鈕--可見) 已產生暫存師資名單
                But_Sub28.Style("display") = "none"
                But_Save28.Style("display") = cst_inline1
                'But_UPLOAD28.Style("display")=cst_inline1 '準備進入上傳
        End Select
    End Sub

    ''' <summary> 下拉選項，首頁資料顯示 </summary>
    Sub cCreate1()
        Call CREATE_NEW_GUID21()  '產生新的GUID 避免記憶體相同 而異常
        'Hid_GUID22.Value=TIMS.GetGUID()
        Hid_NowDate.Value = Get_NowDate()
        hid_chkmsg.Value = "on"
        'create ViewState(vs_dtTaddress)
        Call GetVSAddress()
        '班級-遠距教學
        'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        rbl_DISTANCE = TIMS.GET_DISTANCE(rbl_DISTANCE, 3)
        '空值， 請顯示 整班為實體教學
        'Dim s_DISTANCE_N As String=TIMS.GET_DISTANCE_N(0, "3")
        'lab_DISTANCE.Text=s_DISTANCE_N
        TIMS.sUtl_SetMaxLen(cst_i_ReviseCont_c_max_length, ReviseCont)
        TIMS.sUtl_SetMaxLen(cst_i_Times_c_max_length, txtTimes)

        Dim drP As DataRow = GET_PLANINFO()
        If drP Is Nothing Then Return
        Dim drC As DataRow = GET_CLASSINFO()
        If drC Is Nothing Then Return

        Try
            Call SHOW_PLANPLANINFO(drP, drC) '建立基本資料
            Call SHOW_PLANPLANINFO_REVISE(drP) '建立班級變更資料
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Dim vMsg As String = Common.GetJsString(ex.Message) '.ToString)
            Dim strScript1 As String = ""
            'strScript1 += String.Concat("alert('", vMsg, "');", vbCrLf)
            strScript1 = String.Concat("<script>", "alert('發生錯誤,請重新查詢選取!!\n", vMsg, "');", vbCrLf)
            strScript1 += String.Concat("location.href='TC_05_001.aspx?ID=", TIMS.ClearSQM(Request("ID")), "';", vbCrLf, "</script>")
            Page.RegisterStartupScript("", strScript1)
            Exit Sub
        End Try

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case UCase(rActCheck) 'Cst_cPlan:申請; Cst_cRevise:變更結果
                Case Cst_cPlan  '申請
                Case Cst_cRevise  '變更結果
                    If rPARTREDUC1 <> "Y" Then Call CreateTrainDesc()
            End Select
        End If

        'andy edit
        ChgItem.Attributes("onchange") = "ChangeState('" & sm.UserInfo.TPlanID & "');"
        ChgItem.Attributes("onclick") = "document.getElementById('hid_chkmsg').value='on';"

        But_Sub.Attributes("onclick") = "return Check_Data('" & sm.UserInfo.TPlanID & "');"
        But_Sub28.Attributes("onclick") = "return Check_Data('" & sm.UserInfo.TPlanID & "');"
        EGenSci.Attributes("onchange") = "SetHours();"
        EProSci.Attributes("onchange") = "SetHours();"

        If RIDValue.Value <> "" Then
            Button6.Attributes("onclick") = "wopen('../../Common/TechID.aspx?type=Addx&RID=" & RIDValue.Value & "&TextField=TeacherName1_2&ValueField=NewData11_1&CTName='+document.getElementById('NewData11_1').value,'Tech',700,720,1);"  'Addx(產投任課教師)
            Button6_2.Attributes("onclick") = "wopen('../../Common/TechID.aspx?type=Addy&RID=" & RIDValue.Value & "&TextField=TeacherName2_2&ValueField=NewData20_1&CTName='+document.getElementById('NewData20_1').value,'Tech',700,720,1);"  'Addy(產投助教)
        Else
            Button6.Attributes("onclick") = "alert('該計畫機構資訊有誤!!!');"
            Button6_2.Attributes("onclick") = "alert('該計畫機構資訊有誤!!!');"
        End If

        'by Jimmy 20090520 add 3+2郵遞區號查詢 link
        LitNewData8_1.Text = TIMS.Get_WorkZIPB3Link2()
        'lbZIP6W3=TIMS.Get_WorkZIPB3Link(lbZIP6W3, 3)
        'lbZIP6W2=TIMS.Get_WorkZIPB3Link(lbZIP6W2, 2)

        Dim Button1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(NewData8_1, NewData8_3, hidNewData8_6W, CTName, NewData8_2)
        Button1.Attributes.Add("onclick", Button1_Attr_VAL)

        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            TIMS.Tooltip(PackageTypeNew, "充電起飛計畫不可選擇非包班!!")
            PackageTypeNew.Attributes("onclick") = "GetPackageName54();"
        End If
    End Sub

    ''' <summary>每次重載執行</summary>
    Sub Utl_EveryCreate1()
        'lab_New_Examdate.Text=""
        '70:區域產業據點職業訓練計畫(在職)
        flag_TPlanID70_1 = (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)
        '如果該班並無甄試，則不卡控甄試日期欄位，可為空 (圖四) 'lab_New_Examdate.
        lab_New_Examdate.Text = If(flag_TPlanID70_1, "(若班級無甄試，則甄試日期欄位，可為空)", "")

        'msg1.Text=""
        msg2.Text = ""
        msg3.Text = ""
        msg4.Text = ""
        vsShowmsg4 = ""

        '**by Milor 20080507----end,'Dim sql As String="",'Dim dt As DataTable,'Dim dr As DataRow
        Try
            hidReqID.Value = TIMS.ClearSQM(Request("ID")) 'Request("ID")
            rPlanID = "" & TIMS.ClearSQM(Request("PlanID")) 'Request("PlanID")
            rComIDNO = "" & TIMS.ClearSQM(Request("cid")) 'Request("cid")
            rSeqNo = "" & TIMS.ClearSQM(Request("no")) 'Request("no")
            'sPCS1=rPlanID & "x" & rComIDNO & "x" & rSeqNo
            sPCS1 = String.Format("{0}x{1}x{2}", rPlanID, rComIDNO, rSeqNo)

            rActCheck = "" & UCase(TIMS.ClearSQM(Request("check"))) 'PLAN_PLANINFO/PLAN_REVISE
            rSCDate = "" & TIMS.ClearSQM(Request("CDate"))
            If rSCDate <> "" Then
                If Not TIMS.IsDate1(rSCDate) Then rSCDate = ""
                Hid_rCDATE.Value = TIMS.Cdate3(rSCDate)
            End If
            iSubSeqNO = If(TIMS.ClearSQM(Request("SubNo")) <> "", TIMS.CINT1(TIMS.ClearSQM(Request("SubNo"))), 0)
            'If rSubSeqNO <> "" AndAlso Not IsNumeric(rSubSeqNO) Then rSubSeqNO=""
            'AltDataID=11 (INT)
            'Dim rAltDataID As String
            rAltDataID = TIMS.ClearSQM(Request("AltDataID"))
            If rAltDataID <> "" Then
                'Common.SetListItem(ChgItem, rAltDataID)
                rAltDataID = TIMS.CINT1(rAltDataID)
                chgState.Value = rAltDataID
            End If
            rPARTREDUC1 = TIMS.ClearSQM(Request("PARTREDUC1"))

            hidReqPlanID.Value = rPlanID ' TIMS.ClearSQM(Request("PlanID")) 'Request("PlanID") planid
            hidReqcid.Value = rComIDNO ' TIMS.ClearSQM(Request("cid")) 'Request("cid") comidno
            hidReqno.Value = rSeqNo ' TIMS.ClearSQM(Request("no")) 'Request("no") seqno
            hidReqcheck.Value = rActCheck 'UCase(TIMS.ClearSQM(Request("check"))) 'Request("check")
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Dim vMsg As String = Common.GetJsString(ex.Message) '.ToString)
            Dim strScript1 As String = ""
            'strScript1="<script>" 'strScript1 += "alert('" & vMsg & "');" + vbCrLf
            strScript1 = String.Concat("<script>", "alert('發生錯誤,請重新查詢選取!\n", vMsg, "');", vbCrLf)
            strScript1 += String.Concat("location.href='TC_05_001.aspx?ID=", TIMS.ClearSQM(Request("ID")), "';", vbCrLf, "</script>")
            Page.RegisterStartupScript("", strScript1)
            Exit Sub
        End Try

        'Dim TD_1 As HtmlTableCell=FindControl("TD_1")
        'Dim TD_2 As HtmlTableCell=FindControl("TD_2")
        If ViewState(vs_UpdateTrainDesc) <> "N" Then
            TD_1.Style("display") = "none"
            TD_2.Style("display") = "none"
        End If
        hid_MaxChgItem.Value = Cst_iMaxChgItem
        hid_TPlanID28AppPlan.Value = TIMS.Cst_TPlanID28AppPlan

        'TIMS
        But_Sub.Style("display") = cst_inline1 '一般計畫儲存
        But_Sub28.Style("display") = "none" '產投計畫確定
        But_Save28.Style("display") = "none" '產投計畫儲存
        'But_UPLOAD28.Style("display")="none" '產投計畫儲存(上傳檔案)

        ltlInserNextFlag.Text = "0"  '預設關掉是否繼續下一筆的判斷

        '是否顯示
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            But_Sub.Style("display") = "none"
            Select Case UCase(rActCheck) 'Cst_cPlan:申請; Cst_cRevise:變更結果
                Case Cst_cPlan  '申請
                    Utl_CanSave_ButSub28()

                Case Cst_cRevise  '變更結果
                    But_Sub28.Style("display") = "none" '儲存鈕
                    But_Save28.Style("display") = "none" '確定鈕
                    'But_UPLOAD28.Style("display")="none" '產投計畫儲存(上傳檔案)
                    If rPARTREDUC1 = "Y" Then Utl_CanSave_ButSub28()

            End Select
        End If

        ''產投使用／遠距教學 暫不啟用
        'flag_StopDISTANCE=If(TIMS.Utl_GetConfigSet("STOP_DISTANCE").Equals("Y"), True, False)
        dt_KEY_COSTITEM = TIMS.GET_KEY_COSTITEMdt1(gobjconn)
    End Sub

    ''' <summary> 變更項目設計 </summary>
    Sub SHOW_CHGITEM_1(ByVal ChgStateVal As String)

#Region "(No Use)"
        '<asp@ListItem Value="1">訓練期間</asp@ListItem>
        '<asp@ListItem Value="2">訓練時段</asp@ListItem>
        '<asp@ListItem Value="3">訓練課程地點</asp@ListItem>
        '<asp@ListItem Value="4">課程編配</asp@ListItem>
        '<asp@ListItem Value="5">訓練師資</asp@ListItem>
        '<asp@ListItem Value="6">班別名稱</asp@ListItem>
        '<asp@ListItem Value="7">期別</asp@ListItem>
        '<asp@ListItem Value="8">上課地址</asp@ListItem>
        '<asp@ListItem Value="9">停辦</asp@ListItem>
        '<asp@ListItem Value="10">上課時段</asp@ListItem>
        '<asp@ListItem Value="11">師資</asp@ListItem>
        '<asp@ListItem Value="12">招生人數</asp@ListItem>
        '<asp@ListItem Value="13">增班</asp@ListItem>
        '<asp@ListItem Value="14">學(術)科場地</asp@ListItem>
        '<asp@ListItem Value="15">上課時間</asp@ListItem>

        '<asp:ListItem Value="0" >== 請選擇 ==</asp:ListItem>
        '<asp:ListItem Value="1" > 訓練期間</asp: ListItem>
        '<asp:ListItem Value="2">訓練時段(課程互換)</asp:ListItem>
        '<asp:ListItem Value="3" > 訓練課程地點</asp: ListItem>
        '<asp:ListItem Value="4">課程編配</asp:ListItem>
        '<asp:ListItem Value="5" > 訓練師資</asp: ListItem>
        '<asp:ListItem Value="6">班別名稱</asp:ListItem>
        '<asp:ListItem Value="7" > 期別</asp: ListItem>
        '<asp:ListItem Value="8">上課地址</asp:ListItem>
        '<asp:ListItem Value="9" > 停辦</asp: ListItem>
        '<asp:ListItem Value="10">上課時段</asp:ListItem>
        '<asp:ListItem Value="11" > 師資</asp: ListItem>
        '<asp:ListItem Value="20">助教</asp:ListItem>
        '<asp:ListItem Value="12" > 核定人數</asp: ListItem>
        '<asp:ListItem Value="13">增班</asp:ListItem>
        '<asp:ListItem Value="14" > 學(術)科場地</asp: ListItem>
        '<asp:ListItem Value="15">上課時間</asp:ListItem>
        '<asp:ListItem Value="18" > 課程表</asp: ListItem>
        '<asp:ListItem Value="17">報名日期</asp:ListItem>
        '<asp:ListItem Value="19" > 包班種類</asp: ListItem>
        '<asp:ListItem Value="21">訓練費用</asp:ListItem>
        '<asp:ListItem Value="16" > 其他</asp: ListItem>
#End Region

        '調整可變更內容-- --Start
        Dim ChgItemSortVal As String = "1,2,3,4,5,6,7,8,9,10,11,20,12,13,14,15,18,22,17,19,21,16"
        Dim ChgItemName As String() '將變更項目名稱定義到陣列之中
        '**by Milor 20080507--將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改----start
        '2008-05-21 andy 新增「課程表」
        'ChgItemName=New String() {"開、結訓日期", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別", "期別", "上課地址", "停辦", "上課時段", "師資", "核定人數", "增班", "上課地點", "上課時間", "其他", "報名日期", "課程表", "包班種類"}
        'ChgItemName=New String() {"訓練期間", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別名稱", "期別", "上課地址", "申請停辦", "上課時段", "師資", "核定人數", "增班", "學(術)科場地", "上課時間", "其他", "報名日期", "包班種類"}
        'DISTANCE '遠距教學 BY AMU 20210610

        '將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改
        '產學訓套用的顯示字串  / '非產學訓套用的顯示字串
        ChgItemName = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.TPlanID28ChgItemName, TIMS.TPlanIDChgItemName)

        Dim ChgItemSortAry As String() = ChgItemSortVal.Split(",")
        With ChgItem.Items
            .Clear() '清理
            .Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, "0")) '請選擇
            For i_CI As Integer = 0 To ChgItemSortAry.Length - 1
                Dim str_val As String = ChgItemSortAry(i_CI)
                Dim str_txt As String = ChgItemName(TIMS.CINT1(str_val) - 1)
                .Add(New ListItem(str_txt, str_val))
            Next
        End With

        'Cst_i包班種類
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '未選擇 包班種類 'true:(未使用)未選擇 包班種類(移除) 'false:(使用)選擇 包班種類(保留)
            If flag_PackageType_NOUSE Then
                ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類)) '未選擇 包班種類
            End If
        End If

        '產投使用／遠距教學 暫不啟用
        flag_StopDISTANCE = If(TIMS.Utl_GetConfigSet("STOP_DISTANCE").Equals("Y"), True, False)

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用／產投／充電 自辦職前
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練時段))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練地點))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i課程編配))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練師資))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i班別名稱))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課時段))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i增班))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i報名日期)) '20081015 andy 報名日期產投無此項目
            '**by Milor 20080507--97年產學訓只剩下1.開、結訓日期；9.停辦；11.師資；14.上課地點；15.上課時間；16.其他----start
            'If sm.UserInfo.Years >= 2008 Then
            'End If
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i期別)) '2008停用 期別:7
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課地址)) '2008停用 上課地址:8
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i核定人數)) '201111 '開放因應無薪假再出發 '201203'有配套措失，手冊沒有此功能故移除
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練費用)) '產投 暫無訓練費用 
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i遠距教學)) '產投使用(目前已無疫情，爰請將變更項目「遠距教學」隱藏)

            '20080722 andy edit  暫開放其它項目，因課程表變更尚未上線
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i其他))
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類))
            '**by Milor 20080507----end
            '產業人才投資方案的上課時間／時段問題
            '2008/1/24 由 上課時間 改為 上課時段 '2008/4/24再改回 上課時間 以後將不再改為上課時段 by 豪哥/AMU
            'ChgItem.Items.FindByValue("15").Text="上課時段"
            '產投使用／遠距教學 暫不啟用
            If flag_StopDISTANCE Then ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i遠距教學))
        Else
            '一般TIMS計畫 ，非產投類
            '未開班把不能變更狀態的移除
            If ViewState(vs_OCID) = "" Then
                ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練地點))
                ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i訓練師資))
            End If
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課時段))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i核定人數))
            '經PM發現 TIMS計劃應該無此功能(停辦:9、師資:11、增班:13)-  2008-01-03 by AMU
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue("9"))  '開放停辦功能
            '經PM發現一般TIMS計劃應該無此功能(停辦)-  2008-01-03 by  Andy FindByValue("9"))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i停辦)) ''停辦
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i師資))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i助教)) '職前移除
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i增班)) ''增班
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i科場地))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i上課時間))
            'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i其他)) '20080923 一般計畫其它項目保留
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i報名日期))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i課程表))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類))
            ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i遠距教學)) '產投使用／在職暫不用
        End If
        '調整可變更內容 --End

        Common.SetListItem(ChgItem, ChgStateVal)
    End Sub

    ''' <summary>建立基本資料-CLASS_CLASSINFO</summary>
    ''' <returns></returns>
    Function GET_CLASSINFO() As DataRow
        'parms.Clear()
        Dim parms As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNo}}
        Dim sql As String = ""
        'CLASS_CLASSINFO d 'sql &= " SELECT convert(varchar, a.STDate-1, 111) SignUpEDate" & vbCrLf
        sql &= " SELECT convert(varchar, dateadd(day,-1,a.STDate), 111) SignUpEDate" & vbCrLf
        sql &= " ,d.SEnterDate" & vbCrLf 'date
        sql &= " ,d.FEnterDate" & vbCrLf 'date
        sql &= " ,d.EXAMDATE" & vbCrLf 'date
        sql &= " ,d.ExamPeriod" & vbCrLf
        sql &= " ,d.CheckInDate" & vbCrLf 'date
        sql &= " ,b.TPLANID" & vbCrLf '/*PK*/ 
        sql &= " ,b.PLANNAME" & vbCrLf
        sql &= " ,b.ISONLINE" & vbCrLf
        sql &= " ,b.PLANTYPE" & vbCrLf
        sql &= " ,b.PLANSNAME" & vbCrLf
        sql &= " ,b.CLSYEAR" & vbCrLf
        sql &= " ,b.PUBPRINT" & vbCrLf
        sql &= " ,b.QUESID" & vbCrLf
        sql &= " ,b.EMAILSEND" & vbCrLf
        sql &= " ,b.REUSABLE" & vbCrLf
        sql &= " ,b.BLACKLIST" & vbCrLf
        sql &= " ,b.QUERYDISPLAY" & vbCrLf
        sql &= " ,b.USEECFA" & vbCrLf
        sql &= " ,b.PROPERTYID" & vbCrLf
        sql &= " ,b.USECFIRE1" & vbCrLf
        sql &= " ,a.CJOB_UNKEY" & vbCrLf
        'sql &= ",s.Cjob_NO" & vbCrLf
        'sql &= ",s.Cjob_Name" & vbCrLf
        sql &= " ,CASE WHEN c.JobID IS NULL THEN c.TrainID ELSE c.JobID END TrainID" & vbCrLf
        sql &= " ,CASE WHEN c.JobID IS NULL THEN c.trainName ELSE c.JobName END TrainName" & vbCrLf
        sql &= " ,CASE WHEN c.JobID IS NULL THEN c.TrainID ELSE c.JobID END JobID" & vbCrLf
        sql &= " ,CASE WHEN c.JobID IS NULL THEN c.trainName ELSE c.JobName END JobName" & vbCrLf
        sql &= " ,d.OCID" & vbCrLf
        sql &= " ,d.NotOpen" & vbCrLf
        sql &= " ,d.TPeriod" & vbCrLf
        sql &= " ,d.CTName" & vbCrLf
        sql &= " ,d.NORID" & vbCrLf
        sql &= " ,d.OtherReason" & vbCrLf
        sql &= " ,f.OrgName" & vbCrLf
        sql &= " ,a.TMID" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN Key_Plan b ON a.TPlanID=b.TPlanID AND a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNo=@SeqNo" & vbCrLf
        sql &= " JOIN Key_TrainType c ON c.TMID=a.TMID" & vbCrLf
        sql &= " JOIN Auth_Relship e ON e.RID=a.RID" & vbCrLf
        sql &= " JOIN Org_OrgInfo f ON f.OrgID=e.OrgID" & vbCrLf
        sql &= " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY=a.CJOB_UNKEY" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO d ON a.PlanID=d.PlanID AND a.ComIDNO=d.ComIDNO AND a.SeqNo=d.SeqNo AND d.IsSuccess='Y'" & vbCrLf

        Dim drC As DataRow = DbAccess.GetOneRow(sql, gobjconn, parms)
        Return drC
    End Function

    ''' <summary>建立基本資料-PLAN_PLANINFO</summary>
    ''' <returns></returns>
    Function GET_PLANINFO() As DataRow
        'parms.Clear()
        Dim parms As Hashtable = New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNo}}
        'PLAN_PLANINFO a
        'CLASS_CLASSINFO d
        Dim sql As String = ""
        sql &= " SELECT a.*" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN Key_Plan b ON a.TPlanID=b.TPlanID" & vbCrLf
        sql &= " JOIN Key_TrainType c ON c.TMID=a.TMID" & vbCrLf
        sql &= " JOIN Auth_Relship e ON e.RID=a.RID" & vbCrLf
        sql &= " JOIN Org_OrgInfo f ON f.OrgID=e.OrgID" & vbCrLf
        sql &= " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY=a.CJOB_UNKEY" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO d ON a.PlanID=d.PlanID AND a.ComIDNO=d.ComIDNO AND a.SeqNo=d.SeqNo AND d.IsSuccess='Y'" & vbCrLf
        sql &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNo=@SeqNo" & vbCrLf

        Dim drP As DataRow = DbAccess.GetOneRow(sql, gobjconn, parms)
        Return drP
    End Function

    ''' <summary>取得一次性的現在日期 (yyyy/MM/dd) 使用 Hid_NowDate '查詢資料時，使用查詢日期</summary>
    ''' <returns></returns>
    Function Get_NowDate() As String
        '若是空值，取得最新值 '查詢資料時，使用查詢日期
        If rActCheck = Cst_cRevise AndAlso Hid_rCDATE.Value <> "" Then Hid_NowDate.Value = TIMS.Cdate3(Hid_rCDATE.Value)
        'If Hid_NowDate.Value="" Then Hid_NowDate.Value=TIMS.cdate3(Now.Date)
        If Not TIMS.IsDate1(Hid_NowDate.Value) Then Hid_NowDate.Value = TIMS.Cdate3(Now.Date)
        'Hid_NowDate.Value=TIMS.cdate3(Hid_NowDate.Value)
        Return Hid_NowDate.Value
    End Function

    ''' <summary>建立基本資料顯示 PLAN_PLANINFO / CLASS_CLASSINFO</summary>
    Sub SHOW_PLANPLANINFO(ByRef drP As DataRow, ByRef drC As DataRow)
        ViewState(vs_OCID) = ""
        Session("REVISE_ONCLASS") = Nothing
        Session("Revise_BusPackage") = Nothing
        lab_REVISEACCT_Name.Text = TIMS.Get_ACCNAME(sm.UserInfo.UserID, gobjconn)

        'PLAN_PLANINFO a
        'CLASS_CLASSINFO d
        'Dim drP As DataRow=GET_PLANINFO()
        If drP Is Nothing Then Return

        '非產投計畫(TIMS)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            'https://jira.turbotech.com.tw/browse/TIMSC-208
            Dim sPlanKind As String = TIMS.Get_PlanKind(Me, gobjconn)
            Dim iPlanKind As Integer = TIMS.CINT1(sPlanKind)
            Dim iCostMode As Integer = TIMS.GetCostMode(Me, gobjconn)
            Dim iAdmPercent As Integer = 0
            If Convert.ToString(drP("AdmPercent")) <> "" Then iAdmPercent = TIMS.CINT1(drP("AdmPercent"))
            Dim iTaxPercent As Integer = 0
            If Convert.ToString(drP("TaxPercent")) <> "" Then iTaxPercent = TIMS.CINT1(drP("TaxPercent"))
            Hid_PlanKind.Value = iPlanKind
            Hid_CostMode.Value = iCostMode
            Hid_AdmPercent.Value = iAdmPercent
            Hid_TaxPercent.Value = iTaxPercent
            Dim htSS As New Hashtable
            TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
            TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
            TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
            TIMS.SetMyValue2(htSS, "iPlanKind", iPlanKind)
            TIMS.SetMyValue2(htSS, "iCostMode", iCostMode)
            TIMS.SetMyValue2(htSS, "iAdmPercent", iAdmPercent)
            TIMS.SetMyValue2(htSS, "iTaxPercent", iTaxPercent)
            Call SHOW_COSTITEM_1(htSS, gobjconn)
            Call SHOW_COSTITEM_2(htSS, gobjconn)
        End If

        Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, gobjconn)
        'CLASS_CLASSINFO d
        'Dim drC As DataRow=GET_CLASSINFO()
        ViewState(vs_OCID) = Convert.ToString(drC("OCID"))
        Select Case UCase(rActCheck)
            Case Cst_cPlan
                ApplyDate.Text = Get_NowDate() 'CDate(Now.Date).ToString("yyyy/MM/dd")
        End Select
        PointYN.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用 '產投
            PointYN.Visible = True
            PointYN.Text = If(Convert.ToString(drP("PointYN")) = "Y", cst_PointYN_學分班, cst_PointYN_非學分班)
            JobText.Text = "[" & drC("JobID") & "]" & drC("JobName")
        Else
            TrainText.Text = "[" & drC("TrainID") & "]" & drC("TrainName") 'TIMS
        End If
        hid_TMID.Value = Convert.ToString(drP("TMID"))
        RIDValue.Value = Convert.ToString(drP("RID"))
        '遠距教學
        'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        Hid_DISTANCE.Value = Convert.ToString(drP("DISTANCE"))
        lab_DISTANCE.Text = TIMS.GET_DISTANCE_N(0, Convert.ToString(drP("DISTANCE")))
        '預設值-無遠距教學
        Common.SetListItem(rbl_DISTANCE, Convert.ToString(drP("DISTANCE")))

#Region "(No Use)"

        'ViewState("RID")=drP("RID").ToString
        'If IsNumeric(drP("CyclType")) Then
        '    If Int(drP("CyclType")) <> 0 Then ClassName.Text += "第" & Int(drP("CyclType")) & "期"
        'End If

#End Region

        YearList.Text = String.Concat(drP("PlanYear"), "年度")
        '申請階段 '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
        Dim s_APPSTAGE2_NM2 As String = If(Convert.ToString(drP("APPSTAGE")) <> "", TIMS.GET_APPSTAGE2_NM2(Convert.ToString(drP("APPSTAGE"))), "")
        labAPPSTAGE.Text = If(s_APPSTAGE2_NM2 <> "", String.Concat("(", s_APPSTAGE2_NM2, ")"), "")
        'If Convert.ToString(drC("Cjob_NO")) <> "" Then CjobNO.Text="[" & drC("Cjob_NO") & "]"
        'If Convert.ToString(drC("Cjob_Name")) <> "" Then CjobName.Text=drC("Cjob_Name")
        CjobName.Text = TIMS.Get_CJOBNAME(dtSCJOB, Convert.ToString(drC("CJOB_UNKEY")))
        OrgName.Text = drC("OrgName").ToString
        'ClassName.Text=drP("ClassName").ToString
        ClassName.Text = TIMS.GET_CLASSNAME(Convert.ToString(drP("ClassName")), Convert.ToString(drP("CyclType")))
        TRange.Text = Common.FormatDate(drP("STDate")) & "~" & Common.FormatDate(drP("FDDate"))
        STDate.Value = Common.FormatDate(drP("STDate"))
        FDDate.Value = Common.FormatDate(drP("FDDate"))
        ClassFlag.Text = If(drC("OCID").ToString = "", "否", "是")
        'TransFlag='N'
        ClassFlag.Text &= If(drP("TransFlag").ToString = "N", "<font color ='red'>(未轉班)</font>", "")

#Region "(No Use)"

        'Select Case rActCheck
        '    Case Cst_cPlan
        '        SearchMode.Text="申請"
        '        If IsDBNull(dr("AppliedResult")) Then
        '            CheckMode.Text="審核中"
        '        ElseIf dr("AppliedResult")="Y" Then
        '            CheckMode.Text="審核通過"
        '        ElseIf dr("AppliedResult")="N" Then
        '            CheckMode.Text="審核不通過"
        '        End If
        '    Case Cst_cRevise
        '        SearchMode.Text="變更結果"
        '        If IsDBNull(dr("ReviseStatus")) Then
        '            CheckMode.Text="審核中"
        '        ElseIf dr("ReviseStatus").ToString="Y" Then
        '            CheckMode.Text="審核通過"
        '        ElseIf dr("ReviseStatus").ToString="N" Then
        '            CheckMode.Text="審核不通過"
        '        End If
        'End Select

        'SubSeqNO=0
        'If Not IsDBNull(dr("SubSeqNO")) Then
        '    SubSeqNO=dr("SubSeqNO")
        'End If
        'SCDate=""
        'If Not IsDBNull(dr("CDate")) Then
        '    Try
        '        SCDate=Replace(Common.FormatNow(dr("CDate")), "-", "/")
        '    Catch ex As Exception
        '    End Try
        'End If

#End Region

        '建立要變更的資料
        ReviseCont.Text = ""
        '20080825 andy  add 報名日期 
        Old_SEnterDate.Text = ""
        If Not IsDBNull(drC("SEnterDate")) AndAlso IsDate(Convert.ToString(drC("SEnterDate"))) Then
            Old_SEnterDate.Text = TIMS.GetDateTime1(drC("SEnterDate"))
        End If
        Old_FEnterDate.Text = ""
        If Not IsDBNull(drC("FEnterDate")) AndAlso IsDate(Convert.ToString(drC("FEnterDate"))) Then
            Old_FEnterDate.Text = TIMS.GetDateTime1(drC("FEnterDate"))
        End If
        SEnterDate.Value = If(IsDBNull(drC("SEnterDate")), "", Common.FormatDate(drC("SEnterDate")))
        FEnterDate.Value = If(IsDBNull(drC("FEnterDate")), "", Common.FormatDate(drC("FEnterDate")))

        HR1.Items.Clear()
        HR2.Items.Clear()
        For i As Integer = 0 To 23
            HR1.Items.Add(New ListItem(i, i))
            HR2.Items.Add(New ListItem(i, i))
        Next
        Common.SetListItem(HR1, 23)
        Common.SetListItem(HR2, 23)
        HR1.SelectedIndex = 0
        HR2.SelectedIndex = 0
        MM1.Items.Clear()
        MM2.Items.Clear()
        For j As Integer = 0 To 59
            MM1.Items.Add(New ListItem(j, j))
            MM2.Items.Add(New ListItem(j, j))
        Next
        Common.SetListItem(MM1, 59)
        Common.SetListItem(MM2, 59)

        MM1.SelectedIndex = 0
        MM2.SelectedIndex = 0
        Select Case UCase(rActCheck)
            Case Cst_cRevise
                ViewState("HR1") = Convert.ToString(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_1", gobjconn))
                ViewState("HR2") = Convert.ToString(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_2", gobjconn))
            Case Else
                ViewState("HR1") = ""
                ViewState("HR2") = ""
        End Select

        If ViewState("HR1") <> "" AndAlso IsDate(Convert.ToString(ViewState("HR1"))) Then
            Common.SetListItem(HR1, CInt(Convert.ToDateTime(ViewState("HR1")).ToString("HH")))
            Common.SetListItem(MM1, CInt(Convert.ToDateTime(ViewState("HR1")).ToString("mm")))
            New_SEnterDate.Text = Convert.ToDateTime(ViewState("HR1")).ToString("yyyy/MM/dd")
        End If
        If ViewState("HR2") <> "" AndAlso IsDate(Convert.ToString(ViewState("HR2"))) Then
            Common.SetListItem(HR2, CInt(Convert.ToDateTime(ViewState("HR2")).ToString("HH")))
            Common.SetListItem(MM2, CInt(Convert.ToDateTime(ViewState("HR2")).ToString("mm")))
            New_FEnterDate.Text = Convert.ToDateTime(ViewState("HR2")).ToString("yyyy/MM/dd")
        End If
        'SignUpEDate.Value=Common.FormatDate(dr("SignUpEDate"), 2)
        SignUpEDate.Value = CDate(drC("SignUpEDate")).ToString("yyyy/MM/dd")

        'If New_SEnterDate.Text="0001/01/01" Then
        If New_SEnterDate.Text = "" Then
            New_SEnterDate.Text = Convert.ToDateTime(drC("SEnterDate")).ToString("yyyy/MM/dd")
            New_FEnterDate.Text = Convert.ToDateTime(drC("FEnterDate")).ToString("yyyy/MM/dd")
            Common.SetListItem(HR1, CInt(Convert.ToDateTime(ViewState("SEnterDate")).ToString("HH")))
            Common.SetListItem(MM1, CInt(Convert.ToDateTime(ViewState("SEnterDate")).ToString("mm")))
            Common.SetListItem(HR2, CInt(Convert.ToDateTime(ViewState("FEnterDate")).ToString("HH")))
            Common.SetListItem(MM2, CInt(Convert.ToDateTime(ViewState("FEnterDate")).ToString("mm")))
        End If

        '(非)產投動作
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            '20081107andy  訓練期間項目 add 報名日期
            '報名日期
            ViewState("Old_SEnterDate2") = "" 'yyyy/MM/dd
            ViewState("Old_FEnterDate2") = "" 'yyyy/MM/dd
            If Convert.ToString(drC("SEnterDate")) <> "" Then
                ViewState("Old_SEnterDate2") = Convert.ToDateTime(drC("SEnterDate")).ToString("yyyy/MM/dd")
                Old_SEnterDate2.Text = TIMS.GetDateTime1(drC("SEnterDate"))
            Else
                Old_SEnterDate2.Text = ""
            End If
            If Convert.ToString(drC("FEnterDate")) <> "" Then
                ViewState("Old_FEnterDate2") = Convert.ToDateTime(drC("FEnterDate")).ToString("yyyy/MM/dd")
                Old_FEnterDate2.Text = TIMS.GetDateTime1(drC("FEnterDate"))
            Else
                Old_FEnterDate2.Text = ""
            End If
            If Convert.ToString(drC("EXAMDATE")) <> "" Then
                Old_Examdate.Text = TIMS.Cdate3(drC("EXAMDATE"))
                New_Examdate.Text = TIMS.Cdate3(drC("EXAMDATE"))
            End If
            If Convert.ToString(drC("CHECKINDATE")) <> "" Then
                Old_CheckInDate.Text = TIMS.Cdate3(drC("CHECKINDATE"))
                New_CheckInDate.Text = TIMS.Cdate3(drC("CHECKINDATE"))
            End If

            Old_ExamPeriod.Text = "" ' TIMS.GET_ExamPeriod(old_ExamPeriod, gobjconn)
            New_ExamPeriod = TIMS.GET_ExamPeriod(New_ExamPeriod, gobjconn)
            If Convert.ToString(drC("ExamPeriod")) <> "" Then
                HidOld_ExamPeriod.Value = Convert.ToString(drC("ExamPeriod")) 'value
                Common.SetListItem(New_ExamPeriod, drC("ExamPeriod")) '新舊值
                TIMS.GetListText(New_ExamPeriod)
                If New_ExamPeriod.SelectedItem.Text <> "" Then Old_ExamPeriod.Text = TIMS.GetListText(New_ExamPeriod) 'New_ExamPeriod.SelectedItem.Text 'text
            End If

            '取得當日的更新資料
            Select Case rActCheck
                Case Cst_cRevise
                    ViewState("New_SEnterDate2") = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_1", gobjconn)
                    ViewState("New_FEnterDate2") = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_2", gobjconn)
                Case Else
                    ViewState("New_SEnterDate2") = ""
                    ViewState("New_FEnterDate2") = ""
            End Select

            If ViewState("New_SEnterDate2") <> "" AndAlso IsDate(ViewState("New_SEnterDate2")) Then
                New_SEnterDate2.Text = Convert.ToDateTime(ViewState("New_SEnterDate2")).ToString("yyyy/MM/dd")
            Else
                New_SEnterDate2.Text = ViewState("Old_SEnterDate2")  '""
            End If

            If ViewState("New_FEnterDate2") <> "" AndAlso IsDate(ViewState("New_FEnterDate2")) Then
                New_FEnterDate2.Text = Convert.ToDateTime(ViewState("New_FEnterDate2")).ToString("yyyy/MM/dd")
            Else
                New_FEnterDate2.Text = ViewState("Old_FEnterDate2") '""
            End If

            HR3.Items.Clear()
            HR4.Items.Clear()
            For i As Integer = 0 To 23
                HR3.Items.Add(New ListItem(i, i))
                HR4.Items.Add(New ListItem(i, i))
            Next
            Common.SetListItem(HR3, 23)
            Common.SetListItem(HR4, 23)
            HR3.SelectedIndex = 0
            HR4.SelectedIndex = 0
            MM3.Items.Clear()
            MM4.Items.Clear()
            For j As Integer = 0 To 59
                MM3.Items.Add(New ListItem(j, j))
                MM4.Items.Add(New ListItem(j, j))
            Next
            Common.SetListItem(MM3, 59)
            Common.SetListItem(MM3, 59)

            MM3.SelectedIndex = 0
            MM4.SelectedIndex = 0
            Select Case rActCheck
                Case Cst_cRevise
                    ViewState("HR3") = Convert.ToString(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_1", gobjconn))
                    ViewState("HR4") = Convert.ToString(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_2", gobjconn))
                Case Else
                    ViewState("HR3") = ""
                    ViewState("HR4") = ""
            End Select
            If ViewState("HR3") <> "" Then
                Common.SetListItem(HR3, CInt(Convert.ToDateTime(ViewState("HR3")).ToString("HH")))
                Common.SetListItem(MM3, CInt(Convert.ToDateTime(ViewState("HR3")).ToString("mm")))
            Else
                If Old_SEnterDate2.Text <> "" Then
                    Common.SetListItem(HR3, CInt(Convert.ToDateTime(Old_SEnterDate2.Text).ToString("HH")))
                    Common.SetListItem(MM3, CInt(Convert.ToDateTime(Old_SEnterDate2.Text).ToString("mm")))
                End If
            End If
            If ViewState("HR4") <> "" Then
                Common.SetListItem(HR4, CInt(Convert.ToDateTime(ViewState("HR4")).ToString("HH")))
                Common.SetListItem(MM4, CInt(Convert.ToDateTime(ViewState("HR4")).ToString("mm")))
            Else
                If Old_FEnterDate2.Text <> "" Then
                    Common.SetListItem(HR4, CInt(Convert.ToDateTime(Old_FEnterDate2.Text).ToString("HH")))
                    Common.SetListItem(MM4, CInt(Convert.ToDateTime(Old_FEnterDate2.Text).ToString("mm")))
                End If
            End If
        End If

        BSDate.Text = Common.FormatDate(drP("STDate"))
        BEDate.Text = Common.FormatDate(drP("FDDate"))

        Select Case rActCheck
            Case Cst_cRevise
                OldData15_1.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "OldData15_1", gobjconn)
                NewData15_1.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "NewData15_1", gobjconn)
                ASDate.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NEWDATA1_1", gobjconn)
                AEDate.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NEWDATA1_2", gobjconn)
            Case Else
                OldData15_1.Text = ""
                NewData15_1.Text = ""
                ASDate.Text = Convert.ToDateTime(drP("STDate")).ToString("yyyy/MM/dd")
                AEDate.Text = Convert.ToDateTime(drP("FDDate")).ToString("yyyy/MM/dd")
        End Select

        If ASDate.Text <> "" Then ASDate.Text = Convert.ToDateTime(ASDate.Text).ToString("yyyy/MM/dd")
        If AEDate.Text <> "" Then AEDate.Text = Convert.ToDateTime(AEDate.Text).ToString("yyyy/MM/dd")

        SGenSci.Text = If(IsDBNull(drP("GenSciHours")), 0, drP("GenSciHours"))
        EGenSci.Text = If(IsDBNull(drP("GenSciHours")), 0, drP("GenSciHours"))

        SProSci.Text = If(IsDBNull(drP("ProSciHours")), 0, drP("ProSciHours"))
        EProSci.Text = If(IsDBNull(drP("ProSciHours")), 0, drP("ProSciHours"))

        SSumSci.Text = Int(SGenSci.Text) + Int(SProSci.Text)
        If IsDBNull(drP("ProTechHours")) Then
            SProTech.Text = 0
            EProTech.Text = 0
        Else
            SProTech.Text = drP("ProTechHours")
            EProTech.Text = drP("ProTechHours")
        End If
        If IsDBNull(drP("OtherHours")) Then
            SOther.Text = 0
            EOther.Text = 0
        Else
            SOther.Text = drP("OtherHours")
            EOther.Text = drP("OtherHours")
        End If
        ESumSci.Text = Int(EGenSci.Text) + Int(EProSci.Text)
        ClassCName.Text = drP("ClassName").ToString
        ClassCName2.Text = drP("ClassName").ToString
        CyclType.Text = TIMS.FmtCyclType(Convert.ToString(drP("CyclType")))
        Select Case rActCheck
            Case Cst_cRevise
                ChangeCyclType.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i期別, "NewData7_1", gobjconn)
            Case Else
                ChangeCyclType.Text = ""
        End Select

        Dim v_oldTAddress As String = ""
        Dim v_oldzipname As String = ""
        Dim v_oldZIP6W As String = ""
        If drP("TaddressZip").ToString <> "" Then
            v_oldZIP6W = If(Convert.ToString(drP("TaddressZIP6W")) <> "", Convert.ToString(drP("TaddressZIP6W")), Convert.ToString(drP("TaddressZip")))
            v_oldzipname = TIMS.Get_ZipName(drP("TaddressZip"), gobjconn)
            v_oldTAddress = String.Format("({0}){1}{2}", v_oldZIP6W, v_oldzipname, drP("TAddress"))
        End If
        OldData8_1.Value = drP("TaddressZip").ToString
        OldData8_3.Value = TIMS.GetZIPCODEB3(v_oldZIP6W)
        OldData8_2.Value = drP("TAddress").ToString
        TAddress.Text = v_oldTAddress

        Dim ZIPF3 As String = "" '前三碼
        Dim ZIPB3 As String = "" '後3碼
        Dim vTAddress As String = "" '上課地址 
        Select Case rActCheck
            Case Cst_cRevise
                ZIPF3 = Convert.ToString(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課地址, "NewData8_1", gobjconn)) '前三碼
                ZIPB3 = Convert.ToString(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課地址, "NewData8_3", gobjconn))  '後3碼
                vTAddress = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課地址, "NewData8_2", gobjconn) '上課地址 
            Case Else
                ZIPF3 = "" '前三碼
                ZIPB3 = ""  '後3碼
                vTAddress = "" '上課地址 
        End Select

        NewData8_1.Value = ZIPF3 'TIMS.AddZero(ZIPF3, 3)  '前三碼
        NewData8_3.Value = ZIPB3 '後3碼
        hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
        CTName.Text = TIMS.GET_FullCCTName(gobjconn, NewData8_1.Value, NewData8_3.Value)
        NewData8_2.Text = vTAddress 'TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課地址, "NewData8_2", gobjconn) '上課地址 

        Select Case Convert.ToString(drC("NotOpen"))
            Case "N"
                OldData9_1.Text = cst_OldData9_1_開辦中 '"開辦中"
                NewData9_1.Enabled = True
            Case "Y"
                OldData9_1.Text = cst_OldData9_1_停辦中 '"停辦中"
                OldData9_1.ForeColor = Color.Red
                NewData9_1.Enabled = False
        End Select

        Dim NewData9_1key As String = "" 'Nothing:無值/Y:停辦
        Select Case rActCheck
            Case Cst_cRevise
                NewData9_1key = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i停辦, "NewData9_1", gobjconn)  'Nothing:無值/Y:停辦
            Case Else
                NewData9_1key = "" 'Nothing:無值/Y:停辦
        End Select
        If NewData9_1key = "Y" Then NewData9_1.Checked = True

#Region "(No Use)"

        '不開班原因
        'TIMS.Get_NotOpenReason(NORID)
        'For i As Integer=0 To Split(dr("NORID").ToString, ",").Length - 1
        '    For j As Integer=0 To NORID.Items.Count - 1
        '        If Split(dr("NORID").ToString, ",")(i)=NORID.Items(i).Value Then NORID.Items(i).Selected=True
        '    Next
        'Next

#End Region

        OldData10_1.Value = drC("TPeriod").ToString
        NewData10_1 = TIMS.GET_HOURRAN(NewData10_1, gobjconn, sm)
        If Not NewData10_1.Items.FindByValue(drC("TPeriod").ToString) Is Nothing Then TrainTime.Text = NewData10_1.Items.FindByValue(drC("TPeriod").ToString).Text

        'parms_T1.Clear()
        Dim parms_T1 As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNo},
            {"OCID", Convert.ToString(drC("OCID"))}}

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投
            If Not IsDBNull(drC("OCID")) Then
                'Dim sWOScript1 As String=""
                'sWOScript1="wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & RIDValue.Value & "&TB1=" & TeacherDesc11.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                'btn_TCTYPEA_11.Attributes("onclick")=sWOScript1
                ''Dim sWOScript1 As String=""
                'sWOScript1="wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & RIDValue.Value & "&TB1=" & TeacherDesc20.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                'btn_TCTYPEB_20.Attributes("onclick")=sWOScript1

                'OldData11_1
                Dim Class_Teacher As DataTable = TIMS.Get_TEACHERdt(parms_T1, "A", gobjconn)
                TeacherName1.Text = ""
                OldData11_1.Value = ""
                For Each dr1 As DataRow In Class_Teacher.Rows
                    TeacherName1.Text &= String.Concat(If(TeacherName1.Text <> "", ",", ""), dr1("TeachCName"))
                    OldData11_1.Value &= String.Concat(If(OldData11_1.Value <> "", ",", ""), dr1("TechID"))
                    If Hid_NewData11_3.Value = "" AndAlso Convert.ToString(dr1("TEACHERDESC")) <> "" Then Hid_NewData11_3.Value = Convert.ToString(dr1("TEACHERDESC"))
                Next
                'TeacherName1_2.Text=""
                'NewData11_1.Value=""
                TeacherName1_2.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i師資, "NewData11_2", gobjconn)
                NewData11_1.Value = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i師資, "NewData11_1", gobjconn)
                Dim vNewData11_3 As String = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i師資, "NewData11_3", gobjconn)
                If vNewData11_3 <> "" Then Hid_NewData11_3.Value = vNewData11_3

                'OldData20_1
                Dim Class_Teacher2 As DataTable = TIMS.Get_TEACHERdt(parms_T1, "B", gobjconn)
                TeacherName2.Text = ""
                OldData20_1.Value = ""
                For Each dr1 As DataRow In Class_Teacher2.Rows
                    TeacherName2.Text &= String.Concat(If(TeacherName2.Text <> "", ",", ""), dr1("TeachCName"))
                    OldData20_1.Value &= String.Concat(If(OldData20_1.Value <> "", ",", ""), dr1("TechID"))
                    If Hid_NewData20_3.Value = "" AndAlso Convert.ToString(dr1("TEACHERDESC")) <> "" Then Hid_NewData20_3.Value = Convert.ToString(dr1("TEACHERDESC"))
                Next
                'TeacherName2_2.Text=""
                'NewData20_1.Value=""
                TeacherName2_2.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i助教, "NewData20_2", gobjconn)
                NewData20_1.Value = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i助教, "NewData20_1", gobjconn)
                Dim vNewData20_3 As String = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i助教, "NewData20_3", gobjconn)
                If vNewData20_3 <> "" Then Hid_NewData20_3.Value = vNewData20_3
            End If
            Weeks = TIMS.Get_ddlWeeks(Weeks)
        Else
            TeacherName1.Text = drC("CTName").ToString
            TeacherName1.Text = TIMS.Get_CTNAME1(TeacherName1.Text)
        End If
        OldData12_1.Text = drP("TNum").ToString
        Select Case rActCheck
            Case Cst_cRevise
                NewData12_1.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i核定人數, "NewData12_1", gobjconn)
            Case Else
                NewData12_1.Text = ""
        End Select
        OldData13_1.Text = drP("ClassCount").ToString

        'Cst_i包班種類
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If Convert.ToString(drP("PackageType")) <> "" Then
                flag_PackageType_NOUSE = False 'true:(未使用)未選擇 包班種類(移除) 'false:(使用)選擇 包班種類(保留)
                '有選擇可改變咩
                Select Case $"{drP("PackageType")}"
                    Case "1"
                        '非包班
                        'ChgItem.Items.Remove(ChgItem.Items.FindByValue(Cst_i包班種類))
                        flag_PackageType_NOUSE = True 'true:(未使用)未選擇 包班種類(移除) 'false:(使用)選擇 包班種類(保留)
                    Case "2" '2:企業包班
                        Common.SetListItem(PackageTypeNew, $"{drP("PackageType")}")
                        PackageTypeOld.Text = "企業包班"
                    Case "3" '3:聯合企業包班
                        Common.SetListItem(PackageTypeNew, $"{drP("PackageType")}")
                        PackageTypeOld.Text = "聯合企業包班"
                End Select
                hidPackageTypeOld.Value = $"{drP("PackageType")}"
                If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Call CreateBusPackage()
            End If
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用／產投／充電 自辦職前 'list 的 資料 
            NewData14_1b = TIMS.Get_SciPlaceID(NewData14_1b, rComIDNO, 4, "", gobjconn) '學科場地地址1
            NewData14_2b = TIMS.Get_TechPlaceID(NewData14_2b, rComIDNO, 4, "", gobjconn) '術科場地地址1
            NewData14_3 = TIMS.Get_SciPlaceID(NewData14_3, rComIDNO, 4, "", gobjconn) '學科場地地址2
            NewData14_4 = TIMS.Get_TechPlaceID(NewData14_4, rComIDNO, 4, "", gobjconn) '術科場地地址2

            'Hid_NewData8_4.Value=If(Convert.ToString(drP("AddressSciPTID")) <> "", Convert.ToString(drP("AddressSciPTID")), "") '學科場地地址1
            'Hid_NewData8_5.Value=If(Convert.ToString(drP("AddressTechPTID")) <> "", Convert.ToString(drP("AddressTechPTID")), "") '術科場地地址1
            'Hid_NewData8_6.Value=If(Convert.ToString(drP("AddressSciPTID2")) <> "", Convert.ToString(drP("AddressSciPTID2")), "") '學科場地地址2
            'Hid_NewData8_7.Value=If(Convert.ToString(drP("AddressTechPTID2")) <> "", Convert.ToString(drP("AddressTechPTID2")), "") '術科場地地址2

            Button29.Attributes("onclick") = "return CheckAddTime();"
        End If

        '變更項目設計
        Call SHOW_CHGITEM_1(chgState.Value)

    End Sub

    ''' <summary>'建立班級變更資料</summary>
    ''' <param name="drP"></param>
    Sub SHOW_PLANPLANINFO_REVISE(ByRef drP As DataRow)
        'Dim drP As DataRow=GET_PLANINFO()
        If drP Is Nothing Then Return

        Select Case rActCheck
            Case Cst_cPlan
                SearchMode.Text = cst_SearchMode_申請 '"申請"
                Dim s_CheckMode As String = cst_CheckMode_審核中  '"審核中"
                Select Case Convert.ToString(drP("AppliedResult"))
                    Case "Y"
                        s_CheckMode = cst_CheckMode_審核通過 '"審核通過"
                    Case "N"
                        s_CheckMode = cst_CheckMode_審核不通過 '"審核不通過"
                End Select
                CheckMode.Text = s_CheckMode
                labTitle.Text = cst_labTitle_計畫狀態 '"計畫狀態"

                'ApplyDate.Text=Now.Date'往前移動
                bt_clearTech.Attributes.Add("Onclick", "javascript:return ClearLessonTeah();")
                OLessonTeah1.Style("display") = "none"
                OLessonTeah2.Style("display") = "none"
                OLessonTeah3.Style("display") = "none"
                'OLessonTeah1.Attributes.Add("Onclick", "javascript:LessonTeah('Add','' );")
                'OLessonTeah2.Attributes.Add("Onclick", "javascript:LessonTeah('Add2','');")
                OLessonTeah1.Attributes.Add("onClick", "javascript:LessonTeah3('Add','1','OLessonTeah1','OLessonTeah1Value');")
                OLessonTeah2.Attributes.Add("onClick", "javascript:LessonTeah3('Add','2','OLessonTeah2','OLessonTeah2Value');")
                OLessonTeah3.Attributes.Add("onClick", "javascript:LessonTeah3('Add','3','OLessonTeah3','OLessonTeah3Value');")
                '檢查要是有排課，只可以變更結訓日期
                'sql="SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "'"
                'dr=DbAccess.GetOneRow(sql)
                'If Not dr Is Nothing Then
                'ASDate.Text=BSDate.Text
                ' ASDate.ReadOnly=True
                ' IMG1.Visible=False
                ' End If
                Main2_1.Visible = True
                Sub2_1.Visible = False
                Main2_2.Visible = True
                Sub2_2.Visible = False
                Main3_1.Visible = True
                Sub3_1.Visible = False
                Main5_1.Visible = True
                Sub5_1.Visible = False
                Result.Visible = False
                '20100618 andy 預設帶今天,並自動帶入課程內容
                'TimeSDate.Text=Now.Year.ToString() & "/" & Now.Month.ToString() & "/" & Now.Day.ToString()
                'TimeEDate.Text=Now.Year.ToString() & "/" & Now.Month.ToString() & "/" & Now.Day.ToString()
                TimeSDate.Text = TIMS.Cdate3(drP("stdate"))
                TimeEDate.Text = TIMS.Cdate3(drP("fddate"))
                If TimeSDate.Text = "" Then TimeSDate.Text = Now.ToString("yyyy/MM/dd")
                If TimeEDate.Text = "" Then TimeEDate.Text = Now.ToString("yyyy/MM/dd")
                'TimeEDate_TextChanged(sender, e)
                Call ChangeTimeEDate()

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    ApplyDate.Enabled = False
                    imgApplyDate.Visible = False
                    TIMS.Tooltip(ApplyDate, "產投鎖定", True)
                    'OldData14_1.Value=dr("SciPlaceID").ToString
                    'OldData14_2.Value=dr("TechPlaceID").ToString
                    OldData14_1b.Value = Convert.ToString(drP("SciPlaceID"))
                    OldData14_2b.Value = Convert.ToString(drP("TechPlaceID")) 'drP("TechPlaceID").ToString
                    OldData14_3.Value = Convert.ToString(drP("SciPlaceID2")) 'drP("SciPlaceID2").ToString
                    OldData14_4.Value = Convert.ToString(drP("TechPlaceID2")) 'drP("TechPlaceID2").ToString
                    Hid_OldData8_4.Value = If(Convert.ToString(drP("AddressSciPTID")) <> "", Convert.ToString(drP("AddressSciPTID")), "") '學科場地地址1
                    Hid_OldData8_5.Value = If(Convert.ToString(drP("AddressTechPTID")) <> "", Convert.ToString(drP("AddressTechPTID")), "") '術科場地地址1
                    Hid_OldData8_6.Value = If(Convert.ToString(drP("AddressSciPTID2")) <> "", Convert.ToString(drP("AddressSciPTID2")), "") '學科場地地址2
                    Hid_OldData8_7.Value = If(Convert.ToString(drP("AddressTechPTID2")) <> "", Convert.ToString(drP("AddressTechPTID2")), "") '術科場地地址2
                    'If Not NewData14_1.Items.FindByValue(dr("SciPlaceID").ToString) Is Nothing Then SciPlaceID.Text=NewData14_1.Items.FindByValue(dr("SciPlaceID").ToString).Text
                    'If Not NewData14_2.Items.FindByValue(dr("TechPlaceID").ToString) Is Nothing Then TechPlaceID.Text=NewData14_2.Items.FindByValue(dr("TechPlaceID").ToString).Text
                    If Not NewData14_1b.Items.FindByValue(drP("SciPlaceID").ToString) Is Nothing Then SciPlaceIDb.Text = NewData14_1b.Items.FindByValue(drP("SciPlaceID").ToString).Text
                    If Not NewData14_2b.Items.FindByValue(drP("TechPlaceID").ToString) Is Nothing Then TechPlaceIDb.Text = NewData14_2b.Items.FindByValue(drP("TechPlaceID").ToString).Text
                    If Not NewData14_3.Items.FindByValue(drP("SciPlaceID2").ToString) Is Nothing Then SciPlaceID2.Text = NewData14_3.Items.FindByValue(drP("SciPlaceID2").ToString).Text
                    If Not NewData14_4.Items.FindByValue(drP("TechPlaceID2").ToString) Is Nothing Then TechPlaceID2.Text = NewData14_4.Items.FindByValue(drP("TechPlaceID2").ToString).Text

                    Select Case rActCheck
                        Case Cst_cRevise
                            OldData15_1.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "OldData15_1", gobjconn) 'If(IsDBNull(), "", TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "OldData15_1"))
                            NewData15_1.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "NewData15_1", gobjconn) 'If(IsDBNull(TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "NewData15_1", gobjconn)), "", TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, "NewData15_1"))
                        Case Else
                            OldData15_1.Text = ""
                            NewData15_1.Text = ""
                    End Select
                    Call PlanClassTime() '變更前上課時間
                    Call ReviseClassTime(1) '變更後上課時間
                End If

            Case Cst_cRevise
                '建立變更基本資料
                Call SHOW_PLANREVISE()
        End Select

    End Sub

    ''' <summary>建立變更基本資料 PLAN_REVISE</summary>
    Sub SHOW_PLANREVISE()
        labTitle.Text = cst_labTitle_申請狀態

        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", rSCDate) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", iSubSeqNO) 'Request("SubNo")
        'PLAN_REVISE
        Dim dr As DataRow = Get_PlanReviseDataRow(htSS, gobjconn)
        '查無傳入資訊 '基本資料產生問題
        If dr Is Nothing Then Exit Sub

        lab_REVISEACCT_Name.Text = TIMS.Get_ACCNAME(Convert.ToString(dr("REVISEACCT")), gobjconn)

        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
        'Dim flag_PARTREDUC_Y_CanUpdate As Boolean=False
        'If rPARTREDUC1="Y" AndAlso Convert.ToString(dr("PARTREDUC"))="Y" AndAlso Convert.ToString(dr("ReviseStatus"))="" Then flag_PARTREDUC_Y_CanUpdate=True
        Dim flag_PARTREDUC_Y_CanUpdate As Boolean = (rPARTREDUC1 = "Y" AndAlso Convert.ToString(dr("PARTREDUC")) = "Y" AndAlso Convert.ToString(dr("ReviseStatus")) = "")

        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
        Dim flag_PARTREDUC_Enabled As Boolean = flag_PARTREDUC_Y_CanUpdate '(可修改)
        Hid_PARTREDUC_Y_CanUpdate.Value = If(flag_PARTREDUC_Y_CanUpdate, "Y", "")

        ApplyDate.Enabled = False
        imgApplyDate.Visible = False
        TIMS.Tooltip(ApplyDate, "審核鎖定", True)
        ApplyDate.Text = Common.FormatDate(dr("CDate"))

        SearchMode.Text = cst_SearchMode_變更結果 '"變更結果"
        Dim v_ReviseStatus As String = Convert.ToString(dr("ReviseStatus"))
        Dim s_REVISEDATE As String = ""
        Dim s_CheckMode As String = cst_CheckMode_審核中  '"審核中"
        Select Case v_ReviseStatus
            Case "Y"
                s_CheckMode = cst_CheckMode_審核通過 '"審核通過"
                If Convert.ToString(dr("REVISEDATE")) <> "" Then s_REVISEDATE = "(" & TIMS.Cdate3(dr("REVISEDATE")) & ")"
            Case "N"
                s_CheckMode = cst_CheckMode_審核不通過 '"審核不通過"
            Case ""
                s_CheckMode = cst_CheckMode_審核中 '審核中"
            Case Else
                s_CheckMode = TIMS.ClearSQM(v_ReviseStatus)
        End Select
        CheckMode.Text = String.Format("{0}{1}", s_CheckMode, s_REVISEDATE)

        'Common.SetListItem(ChgItem, dr("AltDataID").ToString)
        ChgItem.Enabled = False
        TIMS.Tooltip(ChgItem, "審核鎖定")
        chgState.Value = $"{dr("AltDataID")}" 'rAltDataID

        Select Case $"{dr("AltDataID")}"
            Case Cst_i訓練期間 '開、結訓日期(產投)
                BSDate.Text = dr("OldData1_1")
                BEDate.Text = dr("OldData1_2")
                ASDate.Text = Convert.ToDateTime(dr("NewData1_1")).ToString("yyyy/MM/dd")
                AEDate.Text = Convert.ToDateTime(dr("NewData1_2")).ToString("yyyy/MM/dd")

                IMG1.Visible = If(flag_PARTREDUC_Enabled, True, False)
                IMG2.Visible = If(flag_PARTREDUC_Enabled, True, False) 'False
                Img11.Visible = If(flag_PARTREDUC_Enabled, True, False) 'False
                Img12.Visible = If(flag_PARTREDUC_Enabled, True, False) 'False
                HR3.Enabled = flag_PARTREDUC_Enabled 'False
                HR4.Enabled = flag_PARTREDUC_Enabled 'False
                MM3.Enabled = flag_PARTREDUC_Enabled 'False
                MM4.Enabled = flag_PARTREDUC_Enabled 'False

                Dim MyValueSS As String = ""
                TIMS.SetMyValue(MyValueSS, "rPlanID", rPlanID)
                TIMS.SetMyValue(MyValueSS, "rComIDNO", rComIDNO)
                TIMS.SetMyValue(MyValueSS, "rSeqNo", rSeqNo)

                '非產投
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                    'Call Get_SignUpEDateVal()
                    SignUpEDate.Value = Get_SignUpEDateVal(MyValueSS, gobjconn) 'CDate(vSignUpEDate).ToString("yyyy/MM/dd")
                    Old_SEnterDate2.Text = ""
                    If Not IsDBNull(dr("OldData17_1")) AndAlso IsDate(Convert.ToString(dr("OldData17_1"))) Then Old_SEnterDate2.Text = TIMS.GetDateTime1(dr("OldData17_1"))
                    Old_FEnterDate2.Text = ""
                    If Not IsDBNull(dr("OldData17_2")) AndAlso IsDate(Convert.ToString(dr("OldData17_2"))) Then Old_FEnterDate2.Text = TIMS.GetDateTime1(dr("OldData17_2"))
                    SEnterDate.Value = TIMS.Cdate3(dr("NewData17_1"))
                    FEnterDate.Value = TIMS.Cdate3(dr("NewData17_2"))
                    HR3.Items.Clear()
                    HR4.Items.Clear()
                    For i As Integer = 0 To 23
                        HR3.Items.Add(New ListItem(i, i))
                        HR4.Items.Add(New ListItem(i, i))
                    Next
                    Common.SetListItem(HR3, 23)
                    Common.SetListItem(HR4, 23)
                    'HR3.SelectedIndex=0
                    'HR4.SelectedIndex=0
                    MM3.Items.Clear()
                    MM4.Items.Clear()
                    For j As Integer = 0 To 59
                        MM3.Items.Add(New ListItem(j, j))
                        MM4.Items.Add(New ListItem(j, j))
                    Next
                    Common.SetListItem(MM3, 59)
                    Common.SetListItem(MM4, 59)
                    'MM3.SelectedIndex=0
                    'MM4.SelectedIndex=0
                    New_SEnterDate2.Text = ""
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NewData17_1", gobjconn)
                    New_SEnterDate2.Text = TIMS.Cdate3(MyValue)
                    If MyValue <> "" AndAlso New_SEnterDate2.Text <> "" Then
                        Common.SetListItem(HR3, CInt(Convert.ToDateTime(MyValue).ToString("HH")))
                        Common.SetListItem(MM3, CInt(Convert.ToDateTime(MyValue).ToString("mm")))
                    End If
                    New_FEnterDate2.Text = ""
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NewData17_2", gobjconn)
                    New_FEnterDate2.Text = TIMS.Cdate3(MyValue)
                    If MyValue <> "" AndAlso New_FEnterDate2.Text <> "" Then
                        Common.SetListItem(HR4, CInt(Convert.ToDateTime(MyValue).ToString("HH")))
                        Common.SetListItem(MM4, CInt(Convert.ToDateTime(MyValue).ToString("mm")))
                    End If

                    'New_ExamPeriod.SelectedIndex=-1
                    'OLDDATA3_1, NEWDATA3_1
                    'OLDDATA10_1 NEWDATA10_1
                    New_Examdate.Text = ""
                    Old_Examdate.Text = ""
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "OLDDATA3_1", gobjconn)
                    If MyValue <> "" Then Old_Examdate.Text = TIMS.Cdate3(MyValue)
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NEWDATA3_1", gobjconn)
                    New_Examdate.Text = TIMS.Cdate3(MyValue)
                    'Dim tValue As String=New_ExamPeriod.SelectedValue 'org value
                    'Common.SetListItem(New_ExamPeriod, MyValue) 'TEMP ADP 'set old value (org value)
                    'New_ExamPeriod.SelectedItem.Text 'old text
                    'Common.SetListItem(New_ExamPeriod, tValue) 'TEMP ADP 'org value
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "OLDDATA10_1", gobjconn)
                    If MyValue <> "" Then HidOld_ExamPeriod.Value = MyValue 'old value 
                    If MyValue <> "" Then Old_ExamPeriod.Text = TIMS.GetExamPeriod(MyValue, gobjconn) '轉中文
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NEWDATA10_1", gobjconn)
                    If MyValue <> "" Then Common.SetListItem(New_ExamPeriod, MyValue)

                    'OLDDATA2_1,NEWDATA2_1
                    Old_CheckInDate.Text = ""
                    New_CheckInDate.Text = ""
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "OLDDATA2_1", gobjconn)
                    If MyValue <> "" Then Old_CheckInDate.Text = TIMS.Cdate3(MyValue)
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, "NEWDATA2_1", gobjconn)
                    If MyValue <> "" Then New_CheckInDate.Text = TIMS.Cdate3(MyValue)
                End If

            Case Cst_i訓練時段
                IMG3.Visible = False
                IMG4.Visible = False
                TimeSDate.Text = dr("OldData2_1")
                TimeSClass.Visible = False
                '20080905 andy  edit  修改訓練時段
                EditSClass.Visible = True  '課程名稱
                EditEClass.Visible = True
                EditSClassItem.Visible = True '節次
                EditEClassItem.Visible = True
                TimeEClass.Visible = False
                tb_ClassChg1.Visible = False
                tb_ClassChg2.Visible = False
                TimeSDate.Visible = True
                TimeEDate.Visible = True

                Stime.Text = If(IsDBNull(dr("OldData2_1")), "", dr("OldData2_1"))
                Etime.Text = If(IsDBNull(dr("NewData2_1")), "", dr("NewData2_1"))
                EditSClass.Text = ShowClassList(dr("OldData2_3"), dr("OldData2_2"), "name")
                EditEClass.Text = ShowClassList(dr("NewData2_3"), dr("NewData2_2"), "name")
                EditSClassItem.Text = ShowClassList(dr("OldData2_3"), dr("OldData2_2"), "item")
                EditEClassItem.Text = ShowClassList(dr("NewData2_3"), dr("NewData2_2"), "item")

            Case Cst_i訓練地點
                IMG5.Visible = False
                PlaceDate.Text = dr("OldData3_1")
                SPlace.Visible = False
                EditPlace.Visible = True
                EditPlace.Text = dr("OldData3_2")
                EditPlaceItem.Visible = True
                EditPlaceItem.Text = dr("OldData3_3")
                EPlace.ReadOnly = True
                EPlace.Text = dr("NewData3_1")

            Case Cst_i課程編配
                SSumSci.Text = dr("OldData4_1").ToString
                SGenSci.Text = dr("OldData4_2").ToString
                SProSci.Text = dr("OldData4_3").ToString
                SProTech.Text = dr("OldData4_4").ToString
                SOther.Text = If(IsNumeric(dr("OldData4_5")), dr("OldData4_5"), 0)
                ESumSci.Text = dr("NewData4_1").ToString
                EGenSci.ReadOnly = True
                EGenSci.Text = dr("NewData4_2").ToString
                EProSci.ReadOnly = True
                EProSci.Text = dr("NewData4_3").ToString
                EProTech.ReadOnly = True
                EProTech.Text = dr("NewData4_4").ToString
                EOther.ReadOnly = True
                EOther.Text = If(IsNumeric(dr("NewData4_5")), dr("NewData4_5"), 0)
            Case Cst_i訓練師資 'Cst_i訓練師資
                STeacher.Visible = False 'CheckBoxList
                EditTech.Visible = True '師資
                Dim OldTeach() As String
                OldTeach = Split(dr("OldData5_2"), ",")
                'OldTeach =師資1,助教1,助教2
                'Tech3Label.Visible=True
                Main5_1.Visible = True
                If OldTeach.Length > 1 Then
                    '師資該堂課有2個
                    EditTech.Text = TIMS.Get_TeachCName(OldTeach(0), gobjconn) '師資1
                    'EditTech.Text=TIMS.Get_TeacherName(OldTeach(0))   '師資1
                    Tech2Label.Visible = True
                    EditTech2.Visible = True
                    EditTech2.Text = TIMS.Get_TeachCName(OldTeach(1), gobjconn) '助教1
                    'EditTech2.Text=TIMS.Get_TeacherName(OldTeach(1))   '助教1'師資2
                Else
                    '1個
                    EditTech.Text = TIMS.Get_TeachCName(dr("OldData5_2"), gobjconn) '師資1
                    'EditTech.Text=TIMS.Get_TeacherName(dr("OldData5_2"))
                End If
                If OldTeach.Length > 2 Then
                    '師資該堂課有3個
                    Tech3Label.Visible = True
                    EditTech3.Visible = True
                    EditTech3.Text = TIMS.Get_TeachCName(OldTeach(2), gobjconn) '助教2
                    'EditTech3.Text=TIMS.Get_TeacherName(OldTeach(2))   '助教2'師資2
                End If
                IMG6.Visible = False '日期
                TechDate.Text = Convert.ToString(dr("OldData5_1")) '異動日
                TechDate.Text = TIMS.Cdate3(TechDate.Text)
                EditTechItem.Visible = True
                EditTechItem.Text = Convert.ToString(dr("OldData5_3")) '節次
                Dim NewTeach() As String = Split(dr("NewData5_1"), ",")
                'NewTeach =師資1,助教1,助教2
                OLessonTeah1.ReadOnly = True
                OLessonTeah2.ReadOnly = True '20081124 andy edit
                OLessonTeah3.ReadOnly = True '20151017 BY AMU
                bt_clearTech.Visible = False
                If NewTeach.Length > 1 Then   '師資該堂課有2個
                    OLessonTeah1.Text = TIMS.Get_TeachCName(NewTeach(0), gobjconn) 'TIMS.Get_TeacherName(NewTeach(0))   '師資1
                    OLessonTeah2.Text = TIMS.Get_TeachCName(NewTeach(1), gobjconn) 'TIMS.Get_TeacherName(NewTeach(1))   '依排課作業之規則 師資2=師資1 +12(欄位)
                Else
                    '1個
                    OLessonTeah1.Text = TIMS.Get_TeachCName(dr("NewData5_1"), gobjconn) 'TIMS.Get_TeacherName(dr("NewData5_1"))
                End If
                If NewTeach.Length > 2 Then   '師資該堂課有3個
                    OLessonTeah3.Text = TIMS.Get_TeachCName(NewTeach(2), gobjconn) 'TIMS.Get_TeacherName(NewTeach(2))   '依排課作業之規則 師資3=師資1 +24(欄位)
                End If

            Case Cst_i班別名稱
                ClassCName.Text = dr("OldData6_1")
                ChangeClassCName.Text = dr("NewData6_1")
                If Not flag_PARTREDUC_Y_CanUpdate Then
                    ChangeClassCName.ReadOnly = True
                End If
            Case Cst_i期別
                ClassCName2.Text = dr("OldData6_1")
                CyclType.Text = dr("OldData7_1")
                ChangeCyclType.Text = dr("NewData7_1")
                ChangeCyclType.ReadOnly = True
            Case Cst_i上課地址
                Dim vOldCTName As String = ""
                If Convert.ToString(dr("OldData8_1")) <> "" Then vOldCTName = TIMS.GET_FullCCTName(gobjconn, Convert.ToString(dr("OldData8_1")), Convert.ToString(dr("OldData8_3")))
                TAddress.Text = String.Concat(vOldCTName, dr("OldData8_2"))

                NewData8_1.Value = dr("NewData8_1").ToString
                NewData8_3.Value = dr("NewData8_3").ToString
                hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
                NewData8_2.Text = dr("NewData8_2").ToString
                CTName.Text = TIMS.GET_FullCCTName(gobjconn, NewData8_1.Value, NewData8_3.Value)

                CTName.ReadOnly = True
                NewData8_2.ReadOnly = True
                Button1.Disabled = True

            Case Cst_i停辦 '停辦(產投)
                NewData9_1.Checked = True
            Case Cst_i上課時段
                Common.SetListItem(NewData10_1, dr("NewData10_1"))
                NewData10_1.Enabled = False
            Case Cst_i師資 '師資(產投)
                TeacherName1.Text = dr("OldData11_2").ToString
                OldData11_1.Value = dr("OldData11_1").ToString
                TeacherName1_2.Text = dr("NewData11_2").ToString
                NewData11_1.Value = dr("NewData11_1").ToString
                Hid_NewData11_3.Value = Convert.ToString(dr("NewData11_3"))
                Button6.Visible = If(flag_PARTREDUC_Enabled, True, False)

            Case Cst_i助教 '助教(產投)
                TeacherName2.Text = dr("OldData20_2").ToString
                OldData20_1.Value = dr("OldData20_1").ToString
                TeacherName2_2.Text = dr("NewData20_2").ToString
                NewData20_1.Value = dr("NewData20_1").ToString
                Hid_NewData20_3.Value = Convert.ToString(dr("NewData20_3"))
                Button6_2.Visible = If(flag_PARTREDUC_Enabled, True, False)

            Case Cst_i核定人數
                OldData12_1.Text = dr("OldData12_1").ToString
                NewData12_1.Text = dr("NewData12_1").ToString
            Case Cst_i增班
                OldData13_1.Text = dr("OldData13_1").ToString
                NewData13_1.Text = dr("NewData13_1").ToString

            Case Cst_i科場地 '上課地點(產投)

                OldData14_1b.Value = GET_OldDataVal(Convert.ToString(dr("OldData14_1")), Convert.ToString(dr("NewData14_1")), NewData14_1b, SciPlaceIDb, flag_PARTREDUC_Enabled)
                OldData14_2b.Value = GET_OldDataVal(Convert.ToString(dr("OldData14_2")), Convert.ToString(dr("NewData14_2")), NewData14_2b, TechPlaceIDb, flag_PARTREDUC_Enabled)
                OldData14_3.Value = GET_OldDataVal(Convert.ToString(dr("OldData14_3")), Convert.ToString(dr("NewData14_3")), NewData14_3, SciPlaceID2, flag_PARTREDUC_Enabled)
                OldData14_4.Value = GET_OldDataVal(Convert.ToString(dr("OldData14_4")), Convert.ToString(dr("NewData14_4")), NewData14_4, TechPlaceID2, flag_PARTREDUC_Enabled)

                Hid_OldData8_4.Value = Convert.ToString(dr("OldData8_4")) '學科場地地址
                Hid_OldData8_5.Value = Convert.ToString(dr("OldData8_5")) '術科場地地址
                Hid_OldData8_6.Value = Convert.ToString(dr("OldData8_6")) '學科場地地址2
                Hid_OldData8_7.Value = Convert.ToString(dr("OldData8_7")) '術科場地地址2

                Hid_NewData8_4.Value = Convert.ToString(dr("NewData8_4")) '學科場地地址
                Hid_NewData8_5.Value = Convert.ToString(dr("NewData8_5")) '術科場地地址
                Hid_NewData8_6.Value = Convert.ToString(dr("NewData8_6")) '學科場地地址2
                Hid_NewData8_7.Value = Convert.ToString(dr("NewData8_7"))'術科場地地址2

                'TaddressS2=TIMS.Get_SciPTID(TaddressS2, rComIDNO, 3, gobjconn)
                'TaddressT2=TIMS.Get_TechPTID(TaddressT2, rComIDNO, 3, gobjconn)
                'If IsDBNull(dr("NewData8_4"))=False Then
                '    Common.SetListItem(TaddressS2, dr("NewData8_4").ToString)
                'End If
                'TaddressS2.Enabled=False
                'If IsDBNull(dr("NewData8_5"))=False Then
                '    Common.SetListItem(TaddressT2, dr("NewData8_5").ToString)
                'End If
                'TaddressT2.Enabled=False

            Case Cst_i上課時間 '上課時間(產投)
                Call PlanClassTime() '變更前上課時間
                Call ReviseClassTime("") '變更後上課時間
                Weeks.Enabled = flag_PARTREDUC_Enabled 'False
                txtTimes.Enabled = flag_PARTREDUC_Enabled 'False
                Button29.Enabled = flag_PARTREDUC_Enabled 'False

            Case Cst_i其他 '其他(產投)
                OldData15_1.Text = ""
                NewData15_1.Text = ""
                If Convert.ToString(dr("OldData15_1")) <> "" Then OldData15_1.Text = Replace(dr("OldData15_1"), vbCrLf, "<br>" & vbCrLf)
                If Convert.ToString(dr("NewData15_1")) <> "" Then NewData15_1.Text = Replace(dr("NewData15_1"), vbCrLf, "<br>" & vbCrLf)

            Case Cst_i報名日期
                '20080825 andy  add 報名日期  
                '20081107 andy  add 報名日期
                'Call Get_SignUpEDateVal()
                Dim MyValueSS As String = ""
                TIMS.SetMyValue(MyValueSS, "rPlanID", rPlanID)
                TIMS.SetMyValue(MyValueSS, "rComIDNO", rComIDNO)
                TIMS.SetMyValue(MyValueSS, "rSeqNo", rSeqNo)
                SignUpEDate.Value = Get_SignUpEDateVal(MyValueSS, gobjconn) 'CDate(vSignUpEDate).ToString("yyyy/MM/dd")
                Old_SEnterDate.Text = ""
                If Not IsDBNull(dr("OldData17_1")) Then Old_SEnterDate.Text = TIMS.GetDateTime1(dr("OldData17_1"))
                Old_FEnterDate.Text = ""
                If Not IsDBNull(dr("OldData17_2")) Then Old_FEnterDate.Text = TIMS.GetDateTime1(dr("OldData17_2"))
                SEnterDate.Value = ""
                If Not IsDBNull(dr("OldData17_1")) Then SEnterDate.Value = Common.FormatDate(dr("NewData17_1"))
                FEnterDate.Value = ""
                If Not IsDBNull(dr("OldData17_2")) Then FEnterDate.Value = Common.FormatDate(dr("NewData17_2"))
                HR1.Items.Clear()
                HR2.Items.Clear()
                For i As Integer = 0 To 23
                    HR1.Items.Add(New ListItem(i, i))
                    HR2.Items.Add(New ListItem(i, i))
                Next
                Common.SetListItem(HR1, 23)
                Common.SetListItem(HR2, 23)
                HR1.SelectedIndex = 0
                HR2.SelectedIndex = 0
                MM1.Items.Clear()
                MM2.Items.Clear()
                For j As Integer = 0 To 59
                    MM1.Items.Add(New ListItem(j, j))
                    MM2.Items.Add(New ListItem(j, j))
                Next
                Common.SetListItem(MM1, 59)
                Common.SetListItem(MM2, 59)
                MM1.SelectedIndex = 0
                MM2.SelectedIndex = 0
                New_SEnterDate.Text = ""
                New_FEnterDate.Text = ""
                New_SEnterDate.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_1", gobjconn)
                New_FEnterDate.Text = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_2", gobjconn)
                If New_SEnterDate.Text <> "" AndAlso IsDate(New_SEnterDate.Text) Then
                    New_SEnterDate.Text = Convert.ToDateTime(New_SEnterDate.Text).ToString("yyyy/MM/dd")
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_1", gobjconn)
                    Common.SetListItem(HR1, CInt(Convert.ToDateTime(MyValue).ToString("HH")))
                    Common.SetListItem(MM1, CInt(Convert.ToDateTime(MyValue).ToString("mm")))
                End If
                If New_FEnterDate.Text <> "" AndAlso IsDate(New_FEnterDate.Text) Then
                    New_FEnterDate.Text = Convert.ToDateTime(New_FEnterDate.Text).ToString("yyyy/MM/dd")
                    MyValue = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, "NewData17_2", gobjconn)
                    Common.SetListItem(HR2, CInt(Convert.ToDateTime(MyValue).ToString("HH")))
                    Common.SetListItem(MM2, CInt(Convert.ToDateTime(MyValue).ToString("mm")))
                End If

            Case Cst_i包班種類
                'If $"{dr("NewData4_1")}" <> "" Then Common.SetListItem(PackageTypeNew, $"{dr("NewData4_1")}")
                Common.SetListItem(PackageTypeNew, $"{dr("NewData4_1")}")
                If Not flag_PARTREDUC_Y_CanUpdate Then
                    BusPackageNewHead.Visible = False
                    PackageTypeNew.Enabled = False '(TIMS.Cst_TPlanID54.IndexOf(sm.UserInfo.TPlanID) > -1)
                    txtUname.Enabled = False
                    txtIntaxno.Enabled = False
                    txtUbno.Enabled = False
                    btnAddBusPackage.Enabled = False
                End If

            Case Cst_i訓練費用
                'https://jira.turbotech.com.tw/browse/TIMSC-208
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "rPlanID", rPlanID)
                TIMS.SetMyValue(sCmdArg, "rComIDNO", rComIDNO)
                TIMS.SetMyValue(sCmdArg, "rSeqNo", rSeqNo)
                TIMS.SetMyValue(sCmdArg, "rSCDate", rSCDate)
                TIMS.SetMyValue(sCmdArg, "rSubSeqNO", iSubSeqNO)
                Dim iRCID1 As Integer = TIMS.Get_REVC_RCID(sCmdArg, 1, gobjconn)
                Dim iRCID2 As Integer = TIMS.Get_REVC_RCID(sCmdArg, 2, gobjconn)
                If iRCID1 = 0 AndAlso iRCID1 = 0 Then Exit Select '無資料離開
                Dim dtR1 As DataTable = TIMS.GET_REVISE_COSTITEMdt(sCmdArg, iRCID1, 1, gobjconn)
                Dim dtR2 As DataTable = TIMS.GET_REVISE_COSTITEMdt(sCmdArg, iRCID1, 1, gobjconn)
                Call SHOW_REVISE_COSTITEM(dtR1, iRCID1, 1)
                Call SHOW_REVISE_COSTITEM(dtR2, iRCID2, 2)

            Case Cst_i遠距教學 '遠距教學(產投)
                '遠距教學 'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
                Hid_DISTANCE.Value = Convert.ToString(dr("OldData22_1"))
                lab_DISTANCE.Text = TIMS.GET_DISTANCE_N(0, Convert.ToString(dr("OldData22_1")))
                Common.SetListItem(rbl_DISTANCE, Convert.ToString(dr("NewData22_1")))
                rbl_DISTANCE.Enabled = flag_PARTREDUC_Enabled 'False

        End Select

        If $"{dr("Reason")}" = "" Then
            Result.Visible = False
        Else
            Result.Visible = True
            Reason.Text = Replace($"{dr("Reason")}", vbCrLf, "<BR>")
        End If
        ReviseCont.Text = $"{dr("ReviseCont")}"
        If $"{dr("changeReason")}" <> "" Then Common.SetListItem(changeReason, $"{dr("changeReason")}")

        If Not flag_PARTREDUC_Y_CanUpdate Then
            changeReason.Enabled = False '變更原因
            ReviseCont.ReadOnly = True '變更說明
        End If

        But_Sub.Visible = False '儲存鈕

        Main2_1.Visible = False '原計畫內容
        Sub2_1.Visible = True '原計畫內容 課程 節次

        Main2_2.Visible = False '變更內容
        Sub2_2.Visible = True '變更內容 課程 節次

        Main3_1.Visible = False '原計畫內容
        Sub3_1.Visible = True '原計畫內容 地點 節次
        Sub5_1.Visible = True '原計畫內容 師資
    End Sub

    ''' <summary>
    ''' 顯示必要資料列
    ''' </summary>
    Sub DisplayTR()
        Select Case chgState.Value 'rAltDataID
            Case "0"
            Case Else
                ReviseTable.Style("display") = cst_inline1
                'andy  edit
                '非產投計畫
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then But_Sub.Style("display") = cst_inline1
        End Select

        Select Case TIMS.CINT1(chgState.Value) 'rAltDataID
            Case Cst_i訓練期間 '1 '訓練期間
                TR1_1.Style("display") = cst_inline1
                TR1_2.Style("display") = cst_inline1
                TR1_3.Style("display") = cst_inline1
                '200806026 andy  課程表
                Tr18.Style("display") = cst_inline1 '課程表

                '2008107 Andy  報名起訖 (產學訓不包在內) '非產投計畫
                tb_New_EnterDate2.Style("display") = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1, cst_inline1, "none")
                tb_EnterDate2.Style("display") = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1, cst_inline1, "none")

            Case Cst_i訓練時段 '2  '訓練時段
                TR2_1.Style("display") = cst_inline1
                TR2_2.Style("display") = cst_inline1
            Case Cst_i訓練地點 '3  '訓練課程地點
                TR3_1.Style("display") = cst_inline1
                TR3_2.Style("display") = cst_inline1
            Case Cst_i課程編配 '4  '課程編配
                TR4_1.Style("display") = cst_inline1
                TR4_2.Style("display") = cst_inline1
            Case Cst_i訓練師資 '5  '訓練師資
                TR5_1.Style("display") = cst_inline1
                TR5_2.Style("display") = cst_inline1
            Case Cst_i班別名稱 '6  '班別名稱
                TR6_1.Style("display") = cst_inline1
                TR6_2.Style("display") = cst_inline1
                TR6_3.Style("display") = cst_inline1
            Case Cst_i期別 '7  '期別
                TR7_1.Style("display") = cst_inline1
                TR7_2.Style("display") = cst_inline1
                TR7_3.Style("display") = cst_inline1
                '課程表
                Tr18.Style("display") = cst_inline1

            Case Cst_i上課地址 '8  '上課地址
                TR8_1.Style("display") = cst_inline1
                TR8_2.Style("display") = cst_inline1
                TR8_3.Style("display") = cst_inline1
            Case Cst_i停辦 '9  '申請停辦
                TR9_1.Style("display") = cst_inline1
                TR9_2.Style("display") = cst_inline1
                'TR9_3.Style("display")=cst_inline1
            Case Cst_i上課時段 '10  '上課時段
                TR10_1.Style("display") = cst_inline1
                TR10_2.Style("display") = cst_inline1
            Case Cst_i師資  '11'師資 Cst_i師資 (產投)
                TR11_1.Style("display") = cst_inline1
                TR11_2.Style("display") = cst_inline1
                TR11_3.Style("display") = cst_inline1
                '200806026 andy  課程表
                Tr18.Style("display") = cst_inline1

            Case Cst_i助教 '20  Cst_i助教 (產投)
                TR20_1.Style("display") = cst_inline1
                TR20_2.Style("display") = cst_inline1
                TR20_3.Style("display") = cst_inline1
                '200806026 andy  課程表
                Tr18.Style("display") = cst_inline1

            Case Cst_i核定人數 '12  '招生人數
                TR12_1.Style("display") = cst_inline1
                TR12_2.Style("display") = cst_inline1
            Case Cst_i增班 '13  '增班
                TR13_1.Style("display") = cst_inline1
                TR13_2.Style("display") = cst_inline1
            Case Cst_i科場地 '14  '學(術)科場地
                TR14_1.Style("display") = cst_inline1
                TR14_2.Style("display") = cst_inline1
                '200806026 andy  課程表
                Tr18.Style("display") = cst_inline1

            Case Cst_i上課時間 '15  '上課時間
                TR15_1.Style("display") = cst_inline1
                TR15_2.Style("display") = cst_inline1
                '課程表
                Tr18.Style("display") = cst_inline1

            Case Cst_i其他 '16  '其他
                TR16_1.Style("display") = cst_inline1
                TR16_2.Style("display") = cst_inline1
            Case Cst_i報名日期 '17 '報名日期
                Tr17_1.Style("display") = cst_inline1
                Tr17_2.Style("display") = cst_inline1
            Case Cst_i課程表 '18 '課程表
                '200806026 andy  課程表
                Tr18.Style("display") = cst_inline1

            Case Cst_i包班種類 '19 'Cst_i包班種類
                TR19_1.Style("display") = cst_inline1
                TR19_2.Style("display") = cst_inline1
            Case Cst_i訓練費用 '21@Cst_i訓練費用
                TR21_1.Style("display") = cst_inline1
                TR21_2.Style("display") = cst_inline1
            Case Cst_i遠距教學 '22
                TR22_1.Style("display") = cst_inline1
                TR22_2.Style("display") = cst_inline1
                '課程表
                Tr18.Style("display") = cst_inline1

        End Select

    End Sub

    ''' <summary>
    ''' 隱藏所有資料列
    ''' </summary>
    Sub HidAllTr()
        ReviseTable.Style("display") = "none"
        But_Sub.Style("display") = "none"
        TR1_1.Style("display") = "none"
        TR1_2.Style("display") = "none"
        TR1_3.Style("display") = "none"
        TR2_1.Style("display") = "none"
        TR2_2.Style("display") = "none"
        TR3_1.Style("display") = "none"
        TR3_2.Style("display") = "none"
        TR4_1.Style("display") = "none"
        TR4_2.Style("display") = "none"
        TR5_1.Style("display") = "none"
        TR5_2.Style("display") = "none"
        TR6_1.Style("display") = "none"
        TR6_2.Style("display") = "none"
        TR6_3.Style("display") = "none"
        TR7_1.Style("display") = "none"
        TR7_2.Style("display") = "none"
        TR7_3.Style("display") = "none"
        TR8_1.Style("display") = "none"
        TR8_2.Style("display") = "none"
        TR8_3.Style("display") = "none"
        TR9_1.Style("display") = "none"
        TR9_2.Style("display") = "none"
        'TR9_3.Style("display")="none"
        TR10_1.Style("display") = "none"
        TR10_2.Style("display") = "none"
        TR11_1.Style("display") = "none"
        TR11_2.Style("display") = "none"
        TR11_3.Style("display") = "none"

        TR20_1.Style("display") = "none"
        TR20_2.Style("display") = "none"
        TR20_3.Style("display") = "none"

        TR12_1.Style("display") = "none"
        TR12_2.Style("display") = "none"
        TR13_1.Style("display") = "none"
        TR13_2.Style("display") = "none"
        TR14_1.Style("display") = "none"
        TR14_2.Style("display") = "none"
        'TR14_1b.Style("display")="none"
        'TR14_2b.Style("display")="none"
        TR15_1.Style("display") = "none"
        TR15_2.Style("display") = "none"
        TR16_1.Style("display") = "none"
        TR16_2.Style("display") = "none"
        '2008  andy 報名起迄日    
        Tr17_1.Style("display") = "none"
        Tr17_2.Style("display") = "none"
        'Cst_i包班種類
        TR19_1.Style("display") = "none"
        TR19_2.Style("display") = "none"
        TR21_1.Style("display") = "none"
        TR21_2.Style("display") = "none"
        'Cst_i遠距教學
        TR22_1.Style("display") = "none"
        TR22_2.Style("display") = "none"
        '200806026 andy  課程表
        Tr18.Style("display") = "none"

        '20081107 andy  變更訓練期間時加入報名起訖項目 (產學訓不包在內) 
        tb_New_EnterDate2.Style("display") = "none"
        tb_EnterDate2.Style("display") = "none"
    End Sub

    '建立計畫 企業包班事業單位
    Sub CreateBusPackage()
        DG_BusPackageOld.Visible = False
        DG_BusPackageNew.Visible = False

        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then '非充電計畫者 不可做企業包班新增
            Session("Revise_BusPackage") = Nothing
            Exit Sub
        End If

        Dim parms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNo}}
        Dim sql As String = ""
        sql &= " SELECT uname 企業名稱 ,intaxno 服務單位統一編號 ,ubno 保險證號" & vbCrLf
        sql &= " FROM PLAN_BUSPACKAGE" & vbCrLf
        sql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        sql &= " ORDER BY BPID" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, gobjconn, parms)  '20181008
        If dt1.Rows.Count > 0 Then
            DG_BusPackageOld.Visible = True
            DG_BusPackageOld.DataSource = dt1
            DG_BusPackageOld.DataBind()
        End If

        'If hTPlanID54.Value="" Then '非  '充電起飛計畫  (hTPlanID54.Value="1")
        '    If Not Session("Revise_BusPackage") Is Nothing Then Session("Revise_BusPackage")=Nothing
        '    Exit Function
        'End If
        ''充電起飛計畫 '非 聯合企業包班
        'Select Case PackageType.SelectedValue
        '    Case "3"  '充電起飛計畫' 聯合企業包班
        '        DataGrid4headTable.Visible=True
        '        DataGrid4Table.Visible=True
        '    Case "2"  '充電起飛計畫' 企業包班
        '        DataGrid4headTable.Visible=True
        '        btnAddBusPackage.Visible=False
        '        DataGrid4Table.Visible=False
        '    Case Else
        '        If Not Session("Revise_BusPackage") Is Nothing Then Session("Revise_BusPackage")=Nothing
        '        Exit Function
        'End Select

        Const Cst_PKName As String = "BPID"
        '' Session("Revise_BusPackage")
        'Dim sql As String=""
        Dim dt As DataTable
        'Dim dr As DataRow

        If Session("Revise_BusPackage") Is Nothing Then
            parms.Clear()
            parms.Add("PLANID", rPlanID)
            parms.Add("COMIDNO", rComIDNO)
            parms.Add("SEQNO", rSeqNo)
            Dim Bsql As String = ""
            If rSCDate <> "" AndAlso iSubSeqNO <> 0 Then
                Bsql &= " SELECT * FROM REVISE_BUSPACKAGE" & vbCrLf
                Bsql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SCDATE=@SCDATE" & vbCrLf
                parms.Add("SCDATE", rSCDate)
            Else
                Bsql &= " SELECT * FROM REVISE_BUSPACKAGE" & vbCrLf
                Bsql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SCDATE=@SCDATE" & vbCrLf
                'sql &= "   AND SubSeqNO='1'" & vbCrLf '每天只提供1筆
                parms.Add("SCDATE", ApplyDate.Text)
            End If

            dt = DbAccess.GetDataTable(Bsql, gobjconn, parms)
            dt.Columns(Cst_PKName).AutoIncrement = True
            dt.Columns(Cst_PKName).AutoIncrementSeed = -1
            dt.Columns(Cst_PKName).AutoIncrementStep = -1
            Session("Revise_BusPackage") = dt
        Else
            dt = Session("Revise_BusPackage")
        End If

        'btnAddBusPackage.Visible=False(移到後段動作)
        If rActCheck = Cst_cRevise Then DG_BusPackageNew.Columns(3).Visible = False

        If dt.Rows.Count > 0 Then
            DG_BusPackageNew.Visible = True
            DG_BusPackageNew.DataSource = dt
            DG_BusPackageNew.DataBind()
        End If
        Session("Revise_BusPackage") = dt
    End Sub

    '組合 特殊字串
    Sub ShowCourseList(ByVal SchoolDate As TextBox, ByVal MyCheckBox As CheckBoxList, ByVal msg As Label, Optional ByVal type As String = "")
        'type: Tech:為含教師的顯示 ; "":不含教師的顯示
        'type: room:課程地點顯示

        'Dim dr As DataRow
        Dim sErrmsg As String = ""
        If Not CheckClassSchedule(ViewState(vs_OCID), sErrmsg, gobjconn) Then
            msg.Text = sErrmsg
            Exit Sub
        End If

        '產投無課程資料 / TIMS才有排課資訊
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Exit Sub

        'sql=" select  ocid   from CLASS_SCHEDULE where  OCID=" & ViewState(vs_OCID)
        'dt=DbAccess.GetDataTable(sql)
        'If dt.Rows.Count=0 Then
        '    If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then msg.Text="本班目前尚未排課，請先確認是否已於課程管理完成排課作業！"
        '    Exit Function
        'End If

        If ViewState(vs_OCID) <> "" Then
            'sql=" SELECT Teacher1 ,Teacher2 ,Teacher3 ,Teacher4 ,Teacher5 ,Teacher6 ,Teacher7 ,Teacher8 ,Teacher9 ,Teacher10 ,Teacher11 ,Teacher12, " & _
            '     "         Teacher13 ,Teacher14 ,Teacher15 ,Teacher16 ,Teacher17 ,Teacher18 ,Teacher19 ,Teacher20 ,Teacher21 ,Teacher22 ,Teacher23 ,Teacher24, " & _
            '     "         Class1 ,Class2 ,Class3 ,Class4 ,Class5 ,Class6 ,Class7 ,Class8 ,Class9 ,Class10 ,Class11 ,Class12 " & _
            '     " FROM CLASS_SCHEDULE " & _
            '     " WHERE OCID=" & ViewState(vs_OCID) & " AND SchoolDate='" & SchoolDate.Text & "'"

            Dim parms As Hashtable = New Hashtable From {{"OCID", Convert.ToString(ViewState(vs_OCID))}, {"SCHOOLDATE", SchoolDate.Text}}
            Dim sql As String = ""
            sql &= " SELECT *" & vbCrLf
            sql &= " FROM CLASS_SCHEDULE" & vbCrLf
            sql &= " WHERE OCID=@OCID AND SCHOOLDATE=@SCHOOLDATE"
            'SQL &= "   AND SCHOOLDATE='" & SCHOOLDATE.TEXT & "'"
            Dim dr As DataRow = DbAccess.GetOneRow(sql, gobjconn, parms)

            msg.Text = ""
            If dr Is Nothing Then
                MyCheckBox.Visible = False
                msg.Text = "查無資料(該班的開結訓日期範圍" & TRange.Text & ")"
            Else
                MyCheckBox.Items.Clear()
                'j=0
                'Dim i As Integer=0 '共12節
                Dim j As Integer = 0 '是否為12堂空堂
                Dim j2 As Integer = 0 '是否12堂為無助教1'教師2
                Dim j3 As Integer = 0 '是否12堂為無助教2'教師3
                For i As Integer = 1 To 12 '是否為12堂空堂
                    If Convert.ToString(dr("Class" & i)) = "" Then
                        j += 1
                        j2 += 1
                        j3 += 1
                    Else
                        '20081124 andy edit 
                        'ex:  "2^54273,54278;" --> "TecherID1,(第2筆)TecherID2"  ※ TecherID2=TecherID1+12(欄位)
                        'ex:  "2^54273,54278,54278;" --> "TecherID1,(第2筆)TecherID2,(第3筆)TecherID3"  ※ TecherID2=TecherID1+12(欄位) ※ TecherID2=TecherID1+24(欄位)
                        Dim sText As String = ""
                        Dim sValue As String = ""
                        Dim iType As Integer = 0 '0:異常(沒有師資) 1:只有老師1 ,2:有助教1與老師1 3:助教1,助教2,老師都有
                        Dim clsName As String = TIMS.Get_CourseName(Convert.ToString(dr("Class" & i)), Nothing, gobjconn)
                        Dim romName As String = Convert.ToString(dr("Room" & i))
                        Dim t1Name As String = TIMS.Get_TeachCName(Convert.ToString(dr("Teacher" & i)), gobjconn)
                        Dim t2Name As String = TIMS.Get_TeachCName(Convert.ToString(dr("Teacher" & i + 12)), gobjconn)
                        Dim t3Name As String = TIMS.Get_TeachCName(Convert.ToString(dr("Teacher" & i + 24)), gobjconn)
                        If t3Name <> "" AndAlso t2Name <> "" AndAlso t1Name <> "" Then iType = 3 '有3種師資
                        If t3Name = "" AndAlso t2Name <> "" AndAlso t1Name <> "" Then
                            j3 += 1
                            iType = 2 '有2種師資
                        End If
                        If t3Name = "" AndAlso t2Name = "" AndAlso t1Name <> "" Then
                            j2 += 1
                            j3 += 1
                            iType = 1 '有1種師資
                        End If

                        Select Case type
                            Case "Tech"
                                Select Case iType
                                    Case 3
                                        sText = "第" & i & "節--" & clsName & "--" & t1Name & "," & t2Name & "," & t3Name
                                        sValue = i & "^" & dr("Teacher" & i) & "," & dr("Teacher" & i + 12) & "," & dr("Teacher" & i + 24)
                                    Case 2
                                        sText = "第" & i & "節--" & clsName & "--" & t1Name & "," & t2Name
                                        sValue = i & "^" & dr("Teacher" & i) & "," & dr("Teacher" & i + 12)
                                    Case 1
                                        sText = "第" & i & "節--" & clsName & "--" & t1Name
                                        sValue = i & "^" & dr("Teacher" & i)
                                End Select
                            Case "room"
                                sText = "第" & i & "節--" & clsName & "--" & romName
                                sValue = i & "^" & TIMS.ClearSQM(Convert.ToString(dr("Room" & i)))
                            Case Else 'class
                                Select Case iType
                                    Case 3
                                        sText = "第" & i & "節--" & clsName & "--" & t1Name & "," & t2Name & "," & t3Name
                                    Case 2
                                        sText = "第" & i & "節--" & clsName & "--" & t1Name & "," & t2Name
                                    Case 1
                                        sText = "第" & i & "節--" & clsName & "--" & t1Name
                                End Select
                                sValue = i & "^" & dr("Class" & i)
                        End Select
                        If iType <> 0 AndAlso sText <> "" AndAlso sValue <> "" Then MyCheckBox.Items.Add(New ListItem(sText, sValue))
                    End If
                Next

                hid_NoTechID2.Value = "" '含有無師資2的資料
                hid_NoTechID3.Value = "" '含有無師資2的資料
                MyCheckBox.Visible = True
                If j = 12 Then
                    MyCheckBox.Visible = False
                    msg.Text = "當天無課程資料"
                Else
                    If j2 = 12 Then
                        '無師資2資料'含有無師資2的資料
                        msg.Text = Cst_msg2
                        hid_NoTechID2.Value = "Y"
                    End If
                    If j3 = 12 Then
                        '無師資2資料'含有無師資2的資料
                        If msg.Text <> "" Then msg.Text &= ","
                        msg.Text &= Cst_msg3
                        hid_NoTechID3.Value = "Y"
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TimeSDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimeSDate.TextChanged
        Call ChangeTimeSDate() '變更要異動的開始日
    End Sub

    Private Sub TimeEDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimeEDate.TextChanged
        Call ChangeTimeEDate() '變更要異動的結束日
    End Sub

    '變更要異動的開始日
    Sub ChangeTimeSDate()
        If TimeSDate.Text <> "" Then AddClass_Sch(ViewState(vs_OCID), TimeSDate.Text, msg1)
        If msg1.Text = "" Then
            If ViewState(vs_IsLoaded) = "Y" Then
                SourceLB1.Items.Clear() '清空
                TargetLB1.Items.Clear() '清空
                SourceLB2.Items.Clear() '清空
                TargetLB2.Items.Clear() '清空
                CreateClassList("start", TimeSDate.Text, TimeEDate.Text)
                CreateClassList("end", TimeSDate.Text, TimeEDate.Text)
            Else
                CreateClassList("start", TimeSDate.Text, TimeEDate.Text)
            End If
        Else
            SourceLB1.Items.Clear() '清空
            TargetLB1.Items.Clear() '清空
        End If

        '檢查資料是否已load
        ChkTwoClassLoaded()
        'If TimeSDate.Text=TimeEDate.Text Then   '當欲變更日期為同一天時
        '    If (SourceLB1.Items.Count > 0) Then
        '        Dim classLi As New ListItem
        '        For Each classLi In SourceLB1.Items
        '            TargetLB1.Items.Add(classLi)
        '        Next
        '    End If
        '    SourceLB1.Items.Clear()

        '    btnAdd_1.Visible=False
        '    btnAddAll_1.Visible=False
        '    btnRemove_1.Visible=False
        '    btnRemoveAll_1.Visible=False
        'End If
    End Sub

    '變更要異動的結束日
    Sub ChangeTimeEDate()
        If TimeEDate.Text <> "" AndAlso ViewState(vs_OCID) <> "" Then Call AddClass_Sch(ViewState(vs_OCID), TimeEDate.Text, msg2)
        If msg2.Text = "" Then
            SourceLB1.Items.Clear() '清空
            TargetLB1.Items.Clear() '清空
            SourceLB2.Items.Clear() '清空
            TargetLB2.Items.Clear() '清空
            Call CreateClassList("end", TimeSDate.Text, TimeEDate.Text)
            Call CreateClassList("start", TimeSDate.Text, TimeEDate.Text)
        Else
            SourceLB2.Items.Clear() '清空
            TargetLB2.Items.Clear() '清空
        End If

        '檢查資料是否已load
        ChkTwoClassLoaded()
        'If TimeSDate.Text=TimeEDate.Text Then   '當欲變更日期為同一天時
        '    If (SourceLB1.Items.Count > 0) Then
        '        Dim classLi As New ListItem
        '        For Each classLi In SourceLB1.Items
        '            TargetLB1.Items.Add(classLi)
        '        Next
        '    End If
        '    SourceLB1.Items.Clear()

        '    btnAdd_1.Visible=False
        '    btnAddAll_1.Visible=False
        '    btnRemove_1.Visible=False
        '    btnRemoveAll_1.Visible=False
        'End If
    End Sub

    '檢查資料是否已load
    Private Sub ChkTwoClassLoaded()
        Dim fCanUSE1 As Boolean = (SourceLB1.Items.Count <> 0 AndAlso SourceLB2.Items.Count <> 0)
        btnAdd_1.Visible = fCanUSE1
        btnAddAll_1.Visible = fCanUSE1
        btnRemove_1.Visible = fCanUSE1
        btnRemoveAll_1.Visible = fCanUSE1

        btnAdd_2.Visible = fCanUSE1
        btnAddAll_2.Visible = fCanUSE1
        btnRemove_2.Visible = fCanUSE1
        btnRemoveAll_2.Visible = fCanUSE1
    End Sub

    Private Sub PlaceDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlaceDate.TextChanged
        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Select Case TIMS.CINT1(v_ChgItem)'ChgItem.SelectedValue
            Case Cst_i訓練地點 '"3" '地點
                ShowCourseList(PlaceDate, SPlace, msg3, "room")
            Case Else
                ShowCourseList(PlaceDate, SPlace, msg3)
        End Select
    End Sub

    Private Sub TechDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TechDate.TextChanged
        If vsShowmsg4 = "" Then
            ShowCourseList(TechDate, STeacher, msg4, "Tech")
            vsShowmsg4 = "1"
        End If
    End Sub

    ''' <summary> 類案序號預設為1 ，取得最大序號  (當天)若為同1類，不再增加新序號 ，每1類當天只能申請1次 </summary>
    ''' <param name="v_ChgItem"></param>
    ''' <returns></returns>
    Function GET_MaxSUBSEQNO_28(ByRef v_ChgItem As String) As Integer
        Dim rst_iMAXSubSeqNo As Integer = 1 '0 '(查無所有類案序號預設為1)
        'parms.Clear()
        Dim parms As Hashtable = New Hashtable From {
            {"PlanID", rPlanID},
            {"COMIDNO", rComIDNO},
            {"SEQNO", rSeqNo},
            {"CDATE", TIMS.Cdate2(ApplyDate.Text)},
            {"ALTDATAID", v_ChgItem} 'ChgItem.SelectedValue)
            }
        Dim sqlstr As String = ""
        sqlstr &= " SELECT 'x' FROM PLAN_REVISE" & vbCrLf
        sqlstr &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        sqlstr &= " AND CDATE=@CDATE AND ALTDATAID=@ALTDATAID"

        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, gobjconn, parms)
        If dt.Rows.Count > 0 Then
            '同天使用相同序號
            sqlstr = ""
            sqlstr &= " SELECT MAX(SubSeqNo) MAXSubSeqNo"
            sqlstr &= " FROM PLAN_REVISE" & vbCrLf
            sqlstr &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
            sqlstr &= " AND CDATE=@CDATE" & vbCrLf
            sqlstr &= " AND ALTDATAID=@ALTDATAID"
            parms.Clear()
            parms.Add("PlanID", rPlanID)
            parms.Add("COMIDNO", rComIDNO)
            parms.Add("SEQNO", rSeqNo)
            parms.Add("CDATE", TIMS.Cdate2(ApplyDate.Text))
            parms.Add("ALTDATAID", v_ChgItem) 'ChgItem.SelectedValue)
            rst_iMAXSubSeqNo = DbAccess.ExecuteScalar(sqlstr, gobjconn, parms)
        End If

        If dt.Rows.Count = 0 Then
            '沒有該類案，查詢是否有其他類案
            'sqlstr += " AND CDATE=" & TIMS.to_date(ApplyDate.Text) 'FIX  ORA-01756 QUOTED STRING NOT PROPERLY TERMINATED
            sqlstr = ""
            sqlstr &= " SELECT 'x' FROM PLAN_REVISE" & vbCrLf
            sqlstr &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND CDATE=@CDATE "
            parms.Clear()
            parms.Add("PLANID", rPlanID)
            parms.Add("COMIDNO", rComIDNO)
            parms.Add("SEQNO", rSeqNo)
            parms.Add("CDATE", TIMS.Cdate2(ApplyDate.Text))
            dt = DbAccess.GetDataTable(sqlstr, gobjconn, parms)
            If dt.Rows.Count > 0 Then
                '有其他類案, 序號+1
                sqlstr = ""
                sqlstr &= " SELECT MAX(SubSeqNo)+1 MAXSubSeqNo1 FROM PLAN_REVISE "
                sqlstr &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND CDATE=@CDATE "
                parms.Clear()
                parms.Add("PLANID", rPlanID)
                parms.Add("COMIDNO", rComIDNO)
                parms.Add("SEQNO", rSeqNo)
                parms.Add("CDATE", TIMS.Cdate2(ApplyDate.Text))
                rst_iMAXSubSeqNo = DbAccess.ExecuteScalar(sqlstr, gobjconn, parms)
            End If
        End If

        Return rst_iMAXSubSeqNo
    End Function

    ''' <summary>OJT-21080201：產投-班級變更審核：新增「還原」按鈕</summary>
    ''' <param name="drPR"></param>
    Sub UPDATE_PARTREDUC(ByRef drPR As DataRow)
        If rPARTREDUC1 = "" Then Return
        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
        'Dim flag_PARTREDUC_Y_CanUpdate As Boolean=False
        'If rPARTREDUC1="Y" AndAlso Convert.ToString(drPR("PARTREDUC"))="Y" AndAlso Convert.ToString(drPR("ReviseStatus"))="" Then flag_PARTREDUC_Y_CanUpdate=True
        Dim flag_PARTREDUC_Y_CanUpdate As Boolean = (rPARTREDUC1 = "Y" AndAlso $"{drPR("PARTREDUC")}" = "Y" AndAlso $"{drPR("ReviseStatus")}" = "")
        If flag_PARTREDUC_Y_CanUpdate Then drPR("PARTREDUC") = Convert.DBNull
    End Sub

    ''' <summary>產投儲存</summary>
    Sub Save_Sub_TPlanID28()
        Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)

        Dim strMassage As String = ""
        Dim objstr As String = ""
        Dim objtable As DataTable = Nothing
        Dim objadapter As SqlDataAdapter = Nothing
        Dim objrow As DataRow = Nothing
        Dim TotalItem As String = ""
        Dim TotalItem2 As String = ""
        Dim campare As String = ""
        Dim campare2 As String = ""

        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim iPlanKind As Integer = TIMS.CINT1(TIMS.Get_PlanKind(Me, gobjconn))
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing

        '20080804  andy  start
        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Dim v_changeReason As String = TIMS.GetListValue(changeReason)
        Dim v_PackageTypeNew As String = TIMS.GetListValue(PackageTypeNew)

        'Dim iSubSeqNO As Integer=1 '(查無所有類案序號預設為1)
        iSubSeqNO = GET_MaxSUBSEQNO_28(v_ChgItem)
        If (ReviseCont.Text.Length > 255) Then ReviseCont.Text = TIMS.Get_Substr1(ReviseCont.Text, 255)

        Dim drPP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, gobjconn)
        '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業

        '4個月(特殊開放)
        Dim flag_spec_f1 As Boolean = If(Convert.ToString(drPP("OVERDUE4MTH")) = "Y", True, False) '未發現特殊規則1。 '符合特殊規則規定1(排除4個月限制) 
        'Dim spec_PCSs1 As String="" 'spec_PCSs1=TIMS.Utl_GetConfigSet("spec_PCSs1") '某些班級可使用特殊規則1。
        'If s_SPEC_PCSs1 <> "" AndAlso s_SPEC_PCSs1.IndexOf(sPCS1) > -1 Then flag_spec_f1=True '符合特殊規則規定1(排除4個月限制)
        Dim ck_OCID As String = If(drPP IsNot Nothing, Convert.ToString(drPP("OCID")), "") '已轉班
        Dim v_APPSTAGE As String = If(drPP IsNot Nothing, Convert.ToString(drPP("APPSTAGE")), "") '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業

        'OJT-21080201：產投 -班級變更審核： 新增「還原」按鈕
        'Dim flag_PARTREDUC_Y_CanUpdate As Boolean=False
        'If rPARTREDUC1="Y" AndAlso Convert.ToString(dr("PARTREDUC"))="Y" AndAlso Convert.ToString(dr("ReviseStatus"))="" Then flag_PARTREDUC_Y_CanUpdate=True
        Select Case TIMS.CINT1(v_ChgItem) 'ChgItem.SelectedValue
            Case Cst_i訓練期間
                If ASDate.Text = "" OrElse AEDate.Text = "" Then
                    Common.MessageBox(Me, cst_errmsg_alt4) '"請輸入變更內容的起迄日期!")
                    Exit Sub
                ElseIf Not IsDate(ASDate.Text) OrElse Not IsDate(AEDate.Text) Then
                    Common.MessageBox(Me, cst_errmsg_alt4b) '"請確認輸入變更內容的起迄日期!")
                    Exit Sub
                ElseIf CDate(ASDate.Text) = CDate(BSDate.Text) AndAlso CDate(AEDate.Text) = CDate(BEDate.Text) Then
                    Common.MessageBox(Me, cst_errmsg_alt3) '"新的訓練日期不能與舊日期的相同!")
                    Exit Sub
                End If

                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Dim fg_chk_APPSTAGE As Boolean = (v_APPSTAGE = "1" OrElse v_APPSTAGE = "2" OrElse v_APPSTAGE = "3" OrElse v_APPSTAGE = "4")
                    If Not fg_chk_APPSTAGE Then
                        Common.MessageBox(Me, "申請階段有誤(請確認班級申請資料)!")
                        Exit Sub
                    End If
                    Dim sNOWY As String = Convert.ToString(sm.UserInfo.Years) ' Now.Year
                    Dim sYMD1231 As String = String.Concat(sm.UserInfo.Years, "1231")
                    Dim sYMDNT0301 As String = String.Concat(sm.UserInfo.Years + 1, "0301")
                    Dim sYMDNT0430 As String = String.Concat(sm.UserInfo.Years + 1, "0430")
                    If Not flag_spec_f1 AndAlso v_APPSTAGE = "1" AndAlso ((CDate(ASDate.Text).ToString("MMdd") > "0630") OrElse (CDate(AEDate.Text).ToString("MMdd") > "0831")) Then
                        Common.MessageBox(Me, cst_errmsg_alt61) '申請階段於「上半年」之課程，【開訓日期】不得 > 6/30 ，【結訓日期】必須 <= 8/31
                        Exit Sub
                    ElseIf Not flag_spec_f1 AndAlso v_APPSTAGE = "2" AndAlso ((CDate(ASDate.Text).ToString("yyyyMMdd") > sYMD1231) OrElse (CDate(AEDate.Text).ToString("yyyyMMdd") >= sYMDNT0301)) Then
                        Common.MessageBox(Me, cst_errmsg_alt62) '申請階段於「下半年」之課程，【開訓日期】不得 > 當年度 12/31 ，【結訓日期】必須 <= 隔年2月底
                        Exit Sub
                    ElseIf Not flag_spec_f1 AndAlso v_APPSTAGE = "3" AndAlso (CDate(AEDate.Text).ToString("yyyyMMdd") > sYMDNT0430) Then
                        Common.MessageBox(Me, cst_errmsg_alt63) '申請階段於「政策性產業」之課程，結訓日不得超過翌年 4/30。
                        Exit Sub
                    ElseIf Not flag_spec_f1 AndAlso v_APPSTAGE = "4" AndAlso (CDate(AEDate.Text).ToString("yyyy") <> sNOWY) Then
                        Common.MessageBox(Me, cst_errmsg_alt3b) '【申請階段】為「4:進階政策性產業」時，結訓日期(訓練迄日), 必須在當年度!
                        Exit Sub
                    End If
                End If

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso PointYN.Text.Equals(cst_PointYN_非學分班) Then
                    If ASDate.Text <> "" AndAlso AEDate.Text <> "" Then
                        'AEDate / 訓練起迄日期有誤!
                        If Not IsDate(ASDate.Text) OrElse Not IsDate(AEDate.Text) Then
                            Common.MessageBox(Me, cst_errmsg_alt2) ' "訓練起迄日期有誤!")
                            Exit Sub '離開程式。
                        End If

                        '(1) 開、結訓日期須在 4 個月內。
                        'AEDate / 開訓日期加4個月(若為【非學分班】，訓練起迄日期區間，不得超過4個月)
                        Dim d_tempDate As Date = DateAdd(DateInterval.Month, 4, CDate(ASDate.Text))
                        If Not flag_spec_f1 AndAlso CDate(AEDate.Text) > d_tempDate Then
                            Common.MessageBox(Me, cst_errmsg_alt1) ' "若為【非學分班】，訓練起迄日期區間，不得超過4個月")
                            Exit Sub '離開程式。
                        End If

                        '產投-班級申請，暫時解除開結訓日期不得超過4個月之卡控 by 20210820 (疫情有太多不確定因素)
                        '加卡 所有班級目前暫時結訓日最遲只能到 次年 04/30 (防手殘)， 包含政策性課程 (原邏輯不變)
                        'Dim flag_error_1c As Boolean=False
                        If Not flag_spec_f1 AndAlso ASDate.Text <> "" AndAlso AEDate.Text <> "" AndAlso IsDate(ASDate.Text) AndAlso IsDate(AEDate.Text) Then
                            Dim tmpDate_NY As Date = DateAdd(DateInterval.Year, 1, CDate(ASDate.Text))
                            Dim tempDate As Date = CDate(String.Format("{0}/4/30", tmpDate_NY.Year.ToString()))
                            Dim flag_error_1c As Boolean = (DateDiff(DateInterval.Day, tempDate, CDate(AEDate.Text)) > 0) '"若為【非學分班】，訓練起迄日期區間，迄日不得超過次年04/30" & vbCrLf
                            '若為【非學分班】，訓練起迄日期區間，迄日不得超過次年04/30"
                            If flag_error_1c Then
                                Common.MessageBox(Me, cst_errmsg_alt1c) '
                                Return
                            End If
                        End If

                        'v_APPSTAGE  '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
                        '(2) 申請階段在「上半年」之課程，結訓日期最遲須在當年度8月底前。
                        Dim tempDate2 As Date = CDate(String.Concat(CDate(ASDate.Text).Year, "/8/31"))
                        '(3) 申請階段在「下半年」之課程，結訓日期最遲須在翌年2月底前。
                        Dim tempDate3 As Date = CDate(String.Concat((CDate(ASDate.Text).Year + 1), "/3/1"))
                        If Not flag_spec_f1 AndAlso v_APPSTAGE = "1" AndAlso (DateDiff(DateInterval.Day, tempDate2, CDate(AEDate.Text)) > 0) Then
                            Common.MessageBox(Me, "非學分班,申請階段在「上半年」之課程，結訓日期最遲須在當年度8月底前。")
                            Return
                        ElseIf Not flag_spec_f1 AndAlso v_APPSTAGE = "2" AndAlso (DateDiff(DateInterval.Day, tempDate3, CDate(AEDate.Text)) >= 0) Then
                            Common.MessageBox(Me, "非學分班,申請階段在「下半年」之課程，結訓日期最遲須在翌年2月底前。")
                            Return
                        End If

                        'If CDate(AEDate.Text) > tempDate Then
                        '    Common.MessageBox(Me, "若為【非學分班】，訓練起迄日期區間，不得超過4個月")
                        '    Select Case Convert.ToString(sm.UserInfo.RoleID)
                        '        Case "1", "0" '0	超級使用者1	系統管理者
                        '            '只警告不離開程式。
                        '        Case Else
                        '            Exit Function '離開程式。
                        '    End Select
                        'End If

                    End If
                End If
                'Add by Kevin 2007/07/18 職訓局要求學分班起迄日期必須大於等於12週
                'If sm.UserInfo.TPlanID="28" And PointYN.Text="學分班" Then
                '    If CInt(DateDiff(DateInterval.DayOfYear, CDate(ASDate.Text), CDate(AEDate.Text))) < 77 Then
                '        Common.MessageBox(Me, "學分班訓練日期起迄，必須大於等於12週!")
                '        Exit Sub
                '    End If
                'End If

                'Dim ck_OCID As String=""
                'Dim drPP As DataRow=TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, gobjconn)
                'If drPP IsNot Nothing Then ck_OCID=Convert.ToString(drPP("OCID"))
                '(請記得於本變更申請審核通過後，修正課程表。)
                If ck_OCID <> "" Then
                    Dim dt2M As DataTable = Get_MaxMinDate(ck_OCID, ASDate.Text, AEDate.Text)
                    If dt2M.Rows.Count > 0 Then strMassage = Cst_msg1
                End If

                '20080702 andy 課程表「日期」 -- start 
                Dim ConfirmUpdate As Boolean = False
                'ConfirmUpdate=False
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產學訓修改課程表「日期」 前要先產生暫存的新的訓練日期
                    ViewState(vs_UpdateItemIndex) = Cst_i訓練期間
                    Select Case Session(cst_ss_TEMP1_TrainDescDT)
                        Case Is <> $"{ASDate.Text},{AEDate.Text}"      '暫存新的訓練日期<>目前所選擇新的訓練日期
                            Session(cst_ss_TEMP1_TrainDescDT) = $"{ASDate.Text},{AEDate.Text}"
                            ViewState(vs_UpdateTrainDesc) = "N"
                            Exit Sub
                        Case Is = $"{ASDate.Text},{AEDate.Text}"  '已產生暫存新的訓練日期 
                            Select Case ViewState(vs_UpdateTrainDesc)
                                Case "Y"
                                    ConfirmUpdate = True
                                    Session(cst_ss_TEMP1_TrainDescDT) = ""
                                Case Else
                                    ViewState(vs_UpdateTrainDesc) = "N"
                                    Exit Sub
                            End Select
                        Case Else
                            Session(cst_ss_TEMP1_TrainDescDT) = $"{ASDate.Text},{AEDate.Text}"  '產生暫存新的訓練日期
                            ViewState(vs_UpdateTrainDesc) = "N"
                            Exit Sub
                    End Select

                Else
                    '非產學訓-->直接更新(應該沒有這個狀況)
                    ConfirmUpdate = True
                End If
                If Not ConfirmUpdate Then Exit Sub
                '20080702 andy --end 

                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData1_1") = BSDate.Text
                objrow("OldData1_2") = BEDate.Text
                objrow("NewData1_1") = ASDate.Text
                objrow("NewData1_2") = AEDate.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then '自辦
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("STDate") = CDate(ASDate.Text).ToString("yyyy/MM/dd")
                        dr("FDDate") = CDate(AEDate.Text).ToString("yyyy/MM/dd")
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("STDate") = CDate(ASDate.Text).ToString("yyyy/MM/dd")
                        dr("FDDate") = CDate(AEDate.Text).ToString("yyyy/MM/dd")
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    'Dim i As Integer
                    sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        For i As Integer = 0 To dt.Rows.Count - 1
                            dr = dt.Rows(i)
                            dr("OpenDate") = CDate(ASDate.Text).ToString("yyyy/MM/dd")
                            dr("CloseDate") = CDate(AEDate.Text).ToString("yyyy/MM/dd")
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da)
                        Next
                    End If
                    '更新課程資料
                    'Dim myThreadDelegate As New Threading.ThreadStart(AddressOf SetClassSchedule)
                    'Dim myThread As New Threading.Thread(myThreadDelegate)
                    'myThread.Start()
                    sql = " UPDATE STUD_DATALID SET ResultDate=" & TIMS.To_date(AEDate.Text) & " WHERE OCID='" & ViewState(vs_OCID) & "'"
                    DbAccess.ExecuteNonQuery(sql, gobjconn)
                End If
            Case Cst_i訓練時段
                'Dim i As Integer 'Dim str, str1 As String
                Dim strTimeSClass As String = ""
                For i As Integer = 0 To TimeSClass.Items.Count - 1
                    If TimeSClass.Items(i).Selected Then strTimeSClass &= String.Concat(If(strTimeSClass <> "", ",", ""), TimeSClass.Items(i).Value)
                Next
                If strTimeSClass = "" Then strTimeSClass = "a"
                Dim Onetmp() As String = Split(strTimeSClass, ",")
                'Onetmp=Split(Left(str, Len(str) - 1), ",")
                'Dim Onetmp2(2) As String
                TotalItem = ""
                For i As Integer = 0 To Onetmp.Length - 1
                    Dim Onetmp2() As String = Split(Onetmp(i), "^")
                    TotalItem += Onetmp2(0) & ","
                    If Onetmp2.Length > 1 Then
                        If i = 0 Then
                            campare = Onetmp2(1)
                        Else
                            If campare <> Onetmp2(1) Then
                                Common.MessageBox(Page, "原計畫內容請選擇相同課程!!!")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                Dim strTimeEClass As String = ""
                For i As Integer = 0 To TimeEClass.Items.Count - 1
                    If TimeEClass.Items(i).Selected Then strTimeEClass &= String.Concat(If(strTimeEClass <> "", ",", ""), TimeEClass.Items(i).Value)
                Next
                If strTimeSClass = "" Then strTimeEClass = "a"
                Dim Sectmp() As String = Split(strTimeEClass, ",")
                'If str1="" Then str1="a" 'Sectmp=Split(Left(str1, Len(str1) - 1), ",")
                'Dim Sectmp2(2) As String
                TotalItem2 = ""
                For i As Integer = 0 To Sectmp.Length - 1
                    Dim Sectmp2() As String = Split(Sectmp(i), "^")
                    TotalItem2 += Sectmp2(0) & ","
                    If Sectmp2.Length > 1 Then
                        If i = 0 Then
                            campare2 = Sectmp2(1)
                        Else
                            If campare2 <> Sectmp2(1) Then
                                Common.MessageBox(Page, "變更計畫內容請選擇相同課程!!!")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                If TotalItem = "," Then  '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "原計畫內容請選擇課程!-6FDB")
                    Exit Sub
                End If
                If TotalItem2 = "," Then '判斷變更計畫是否有選課程
                    Common.MessageBox(Page, "變更計畫內容請選擇課程!!!")
                    Exit Sub
                End If
                If Split(Left(TotalItem, Len(TotalItem) - 1), ",").Length <> Split(Left(TotalItem2, Len(TotalItem2) - 1), ",").Length Then         '判斷原計畫及變更計畫節數是否相等
                    Common.MessageBox(Page, "原計畫及變更計畫課程數不相同!!!")
                    Exit Sub
                End If
                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練時段, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData2_1") = TimeSDate.Text
                objrow("OldData2_2") = campare '課程
                objrow("OldData2_3") = Left(TotalItem, Len(TotalItem) - 1)      '節次，逗號分開
                objrow("NewData2_1") = TimeEDate.Text
                objrow("NewData2_2") = campare2 '課程
                objrow("NewData2_3") = Left(TotalItem2, Len(TotalItem2) - 1)    '節次，逗號分開
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    Dim intDt1 As Integer = 0
                    Dim intDt2 As Integer = 0
                    If dt.Rows.Count > 0 Then
                        intDt1 = dt.Select("SchoolDate='" & TimeSDate.Text & "'").Length
                        intDt2 = dt.Select("SchoolDate='" & TimeEDate.Text & "'").Length
                    End If
                    If dt.Rows.Count > 0 AndAlso intDt1 > 0 AndAlso intDt2 > 0 Then
                        Dim dr1 As DataRow = dt.Select("SchoolDate='" & TimeSDate.Text & "'")(0)
                        Dim dr2 As DataRow = dt.Select("SchoolDate='" & TimeEDate.Text & "'")(0)
                        Dim ClassNum1 As Array = Split(Left(TotalItem, Len(TotalItem) - 1), ",")
                        Dim ClassNum2 As Array = Split(Left(TotalItem2, Len(TotalItem2) - 1), ",")
                        Dim OldCalss As Integer = dr1("Class" & ClassNum1(0))
                        Dim NewCalss As Integer = dr2("Class" & ClassNum2(0))
                        For j As Integer = 0 To ClassNum1.Length - 1
                            dr1("Class" & ClassNum1(j)) = NewCalss
                        Next
                        For k As Integer = 0 To ClassNum2.Length - 1
                            dr2("Class" & ClassNum2(k)) = OldCalss
                        Next
                        dr1("ModifyAcct") = sm.UserInfo.UserID
                        dr1("ModifyDate") = Now()
                        dr2("ModifyAcct") = sm.UserInfo.UserID
                        dr2("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i訓練地點 '(產投)訓練地點
                'Dim i As Integer
                Dim str As String = ""
                For i As Integer = 0 To SPlace.Items.Count - 1
                    If SPlace.Items(i).Selected Then str += SPlace.Items(i).Value.Replace(",", ";") & ","
                Next
                Dim Onetmp() As String
                If str = "" Then str = "a"
                Onetmp = Split(Left(str, Len(str) - 1), ",")
                Dim Onetmp2(2) As String
                TotalItem = ""
                For i As Integer = 0 To Onetmp.Length - 1
                    Onetmp2 = Split(Onetmp(i), "^")
                    TotalItem += Onetmp2(0) & ","
                    If Onetmp2.Length > 1 Then
                        If i = 0 Then
                            campare = Onetmp2(1)
                        Else
                            If campare <> Onetmp2(1) Then
                                Common.MessageBox(Page, "原計畫內容請選擇相同地點!!")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                If TotalItem = "," Then           '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "原計畫內容請選擇課程!-7CA1")
                    Exit Sub
                End If
                If EPlace.Text = "" Then         '判斷是否有填更換地點
                    Common.MessageBox(Page, "請填入要更換地點!!!")
                    Exit Sub
                End If
                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練地點, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData3_1") = PlaceDate.Text
                objrow("OldData3_2") = campare          '地點
                objrow("OldData3_3") = Left(TotalItem, Len(TotalItem) - 1)      '節次，逗號分開
                objrow("NewData3_1") = EPlace.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    'Dim dr As DataRow
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND SchoolDate=" & TIMS.To_date(PlaceDate.Text) & ""
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim ClassNum As Array = Split(Left(TotalItem, Len(TotalItem) - 1), ",")
                        For j As Integer = 0 To ClassNum.Length - 1
                            dr("Room" & ClassNum(j)) = EPlace.Text
                        Next
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If
            Case Cst_i課程編配 '更變訓練時數
                If EGenSci.Text = "" Then   '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "請輸入一般學科!!!")
                    Exit Sub
                End If
                If EProSci.Text = "" Then   '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "請輸入一般術科!!!")
                    Exit Sub
                End If
                If EProTech.Text = "" Then  '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "請輸入術科!!!")
                    Exit Sub
                End If
                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i課程編配, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData4_1") = SSumSci.Text
                objrow("OldData4_2") = SGenSci.Text
                objrow("OldData4_3") = SProSci.Text
                objrow("OldData4_4") = SProTech.Text
                objrow("OldData4_5") = SOther.Text
                objrow("NewData4_1") = ESumSci.Text
                objrow("NewData4_2") = EGenSci.Text
                objrow("NewData4_3") = EProSci.Text
                objrow("NewData4_4") = EProTech.Text
                objrow("NewData4_5") = EOther.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    'Dim dr As DataRow
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("GenSciHours") = EGenSci.Text
                        dr("ProSciHours") = EProSci.Text
                        dr("ProTechHours") = EProTech.Text
                        dr("OtherHours") = EOther.Text
                        dr("TotalHours") = TIMS.CINT1(EGenSci.Text) + TIMS.CINT1(EProSci.Text) + TIMS.CINT1(EProTech.Text) + TIMS.CINT1(EOther.Text) '20060525 by Vicient
                        dr("Thours") = TIMS.CINT1(EGenSci.Text) + TIMS.CINT1(EProSci.Text) + TIMS.CINT1(EProTech.Text) + TIMS.CINT1(EOther.Text)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)

                        '20060525 by Vicient start
                        sql = ""
                        sql &= " SELECT * FROM CLASS_CLASSINFO "
                        sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                        dt = DbAccess.GetDataTable(sql, da, gobjconn)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("THours") = TIMS.CINT1(EGenSci.Text) + TIMS.CINT1(EProSci.Text) + TIMS.CINT1(EProTech.Text) + TIMS.CINT1(EOther.Text)
                            dr("LastState") = "M" 'M: 修改(最後異動狀態)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da)
                        End If
                        'end
                    End If
                End If

            Case Cst_i訓練師資
                'Dim i As Integer
                Dim str As String = ""
                For i As Integer = 0 To STeacher.Items.Count - 1
                    'AndAlso STeacher.Items(i).Value <> ""
                    If STeacher.Items(i).Selected Then str += STeacher.Items(i).Value & ","
                Next
                Dim Onetmp() As String
                If str = "" Then str = "a"
                Onetmp = Split(Left(str, Len(str) - 1), ",")
                Dim Onetmp2(2) As String
                TotalItem = ""
                For i As Integer = 0 To Onetmp.Length - 1
                    Onetmp2 = Split(Onetmp(i), "^")
                    TotalItem += Onetmp2(0) & ","
                    If Onetmp2.Length > 1 Then
                        If i = 0 Then
                            campare = Onetmp2(1)
                        Else
                            If campare <> Onetmp2(1) Then
                                'Common.MessageBox(Page, "原計畫內容請選擇相同師資!!!")
                                'Exit Sub
                            End If
                        End If
                    End If
                Next
                If TotalItem = "," Then           '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "原計畫內容請選擇師資!!!")
                    Exit Sub
                End If
                If OLessonTeah1Value.Value = "" Then         '判斷是否有填更換地點
                    Common.MessageBox(Page, "請填入要更換師資!!!")
                    Exit Sub
                End If

                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練師資, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData5_1") = TechDate.Text
                objrow("OldData5_2") = campare          '師資
                objrow("OldData5_3") = Left(TotalItem, Len(TotalItem) - 1)      '節次，逗號分開
                objrow("NewData5_1") = OLessonTeah1Value.Value
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    'Dim dr As DataRow
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND SchoolDate=" & TIMS.To_date(TechDate.Text)
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim ClassNum As Array = Split(Left(TotalItem, Len(TotalItem) - 1), ",")
                        For j As Integer = 0 To ClassNum.Length - 1
                            dr("Teacher" & ClassNum(j)) = OLessonTeah1Value.Value
                        Next
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i班別名稱
                If ChangeClassCName.Text = "" Then         '判斷是否有填班別名稱
                    Common.MessageBox(Page, "請填入要更換班別名稱!!!")
                    Exit Sub
                End If
                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i班別名稱, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData6_1") = ClassCName.Text
                objrow("NewData6_1") = ChangeClassCName.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    'Dim dr As DataRow
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("ClassName") = ChangeClassCName.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If

                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("ClassCName") = ChangeClassCName.Text
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If
            Case Cst_i期別
                Dim sPMS As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNO", rSeqNo}}
                objstr = ""
                objstr &= " SELECT a.CyclType ,c.RID"
                objstr &= " FROM PLAN_PLANINFO a "
                objstr &= " JOIN Org_OrgInfo b ON a.ComIDNO=b.ComIDNO"
                objstr &= " JOIN Auth_Relship c ON b.OrgID=c.OrgID"
                objstr &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO"
                objrow = DbAccess.GetOneRow(objstr, gobjconn, sPMS)

                Dim csPMS As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNO", rSeqNo}}
                Dim chkstr As String = "" 'Dim objrow2 As DataRow=Nothing
                chkstr &= " SELECT * FROM PLAN_PLANINFO "
                chkstr &= " WHERE TransFlag='Y' AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO =@SeqNO"
                Dim objrow2 As DataRow = DbAccess.GetOneRow(chkstr, gobjconn, csPMS)

                Dim sPMS3 As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNO", rSeqNo}}
                Dim objstr3 As String = ""
                objstr3 &= " SELECT b.CLSID ,b.YEARS ,c.CLASSID ,b.CYCLTYPE,b.OCID,b.CLASSCNAME" & vbCrLf
                objstr3 &= " FROM PLAN_PLANINFO a" & vbCrLf
                objstr3 &= " JOIN CLASS_CLASSINFO b ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNO=b.SeqNO" & vbCrLf
                objstr3 &= " JOIN ID_Class c ON b.CLSID=c.CLSID" & vbCrLf
                objstr3 &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO"
                Dim objrow3 As DataRow = DbAccess.GetOneRow(objstr3, gobjconn, sPMS3)

                Dim chkstr3 As String = ""
                Dim chkstr4 As String = ""
                If objrow3 IsNot Nothing Then
                    'objrow3() 變更的期別與開班資料中的期別重覆-產投
                    chkstr3 = ""
                    chkstr3 &= " SELECT 'x'"
                    chkstr3 &= " FROM CLASS_CLASSINFO cc"
                    chkstr3 &= " WHERE cc.CLSID='" & objrow3("CLSID") & "'"
                    chkstr3 &= " AND cc.PlanID=" & rPlanID & vbCrLf
                    chkstr3 &= " AND cc.CyclType='" & ChangeCyclType.Text & "'"
                    chkstr3 &= " AND cc.RID='" & objrow("RID") & "'"
                    chkstr3 &= " AND cc.CLASSCNAME='" & objrow3("CLASSCNAME") & "'"
                    chkstr3 &= " AND cc.OCID != '" & objrow3("OCID") & "'"

                    chkstr4 = ""
                    chkstr4 &= " SELECT 'x'"
                    chkstr4 &= " FROM CLASS_STUDENTSOFCLASS cs "
                    chkstr4 &= " JOIN CLASS_CLASSINFO cc ON cc.ocid=cs.ocid "
                    chkstr4 &= " WHERE cc.PLANID=" & rPlanID 'Request("PlanID")  
                    chkstr4 &= " AND cc.OCID ='" & objrow3("OCID") & "'" 'Request("PlanID")  
                End If
                If ChangeCyclType.Text = "" Then         '判斷是否有填期別
                    Common.MessageBox(Page, "請填入要更換期別!!!")
                    Exit Sub
                End If
                If ChangeCyclType.Text.Length <> 2 Then
                    Common.MessageBox(Page, "期別要輸入2位數字才行!!!")
                    Exit Sub
                End If
                If ChangeCyclType.Text = CyclType.Text Then
                    Common.MessageBox(Page, "變更的期別不可和原期別相同!!!")
                    Exit Sub
                End If
                If chkstr3 <> "" AndAlso DbAccess.GetCount(chkstr3, gobjconn) > 0 Then
                    Common.MessageBox(Page, "變更的期別與開班資料中的期別重覆!!!")
                    Exit Sub
                End If
                'Common.MessageBox(Page, "沒有轉班才可改!!!")
                If objrow2 IsNot Nothing AndAlso chkstr4 <> "" AndAlso DbAccess.GetCount(chkstr4, gobjconn) > 0 Then
                    Common.MessageBox(Page, "此計劃班別已有學員資料，不可變更期別!!!")
                    Exit Sub
                End If

                objstr = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i期別, Cst_sql, gobjconn)
                objtable = DbAccess.GetDataTable(objstr, objadapter, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(objtable, objrow, v_ChgItem)

                objrow("OldData6_1") = ClassCName2.Text
                CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
                objrow("OldData7_1") = CyclType.Text
                objrow("NewData7_1") = ChangeCyclType.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(objrow, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    'Dim dr As DataRow
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim vCyclType As String = TIMS.FmtCyclType(ChangeCyclType.Text)
                        dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim vCyclType As String = TIMS.FmtCyclType(ChangeCyclType.Text)
                        dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If
            Case Cst_i上課地址 '變更訓練地點
                'Dim dr As DataRow
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課地址, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData8_1") = If(OldData8_1.Value = "", Convert.DBNull, OldData8_1.Value)
                dr("OldData8_3") = If(OldData8_3.Value = "", Convert.DBNull, OldData8_3.Value)
                dr("OldData8_2") = If(OldData8_2.Value = "", Convert.DBNull, OldData8_2.Value)

                NewData8_1.Value = TIMS.ClearSQM(NewData8_1.Value)
                NewData8_3.Value = TIMS.ClearSQM(NewData8_3.Value)
                hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
                NewData8_2.Text = TIMS.ClearSQM(NewData8_2.Text)

                dr("NewData8_1") = NewData8_1.Value
                dr("NewData8_3") = If(NewData8_3.Value <> "", NewData8_3.Value, Convert.DBNull) 'NewData8_3.Value.Trim '20090520 fix
                dr("NewData8_2") = NewData8_2.Text

                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(objtable, objadapter)

                If iPlanKind = 1 Then
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TaddressZip") = NewData8_1.Value
                        dr("TaddressZIP6W") = If(hidNewData8_6W.Value <> "", hidNewData8_6W.Value, Convert.DBNull)
                        dr("TAddress") = NewData8_2.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If

                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TaddressZip") = NewData8_1.Value
                        dr("TaddressZIP6W") = If(hidNewData8_6W.Value <> "", hidNewData8_6W.Value, Convert.DBNull)
                        dr("TAddress") = NewData8_2.Text
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i停辦 '變更停辦狀態
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i停辦, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData9_1") = "N"
                dr("NewData9_1") = "Y"
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

            Case Cst_i上課時段
                Dim v_NewData10_1 As String = TIMS.GetListValue(NewData10_1)
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課時段, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData10_1") = OldData10_1.Value
                dr("NewData10_1") = v_NewData10_1 'NewData10_1.SelectedValue
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

            Case Cst_i師資 '產投儲存
                Dim strErrmsg As String = ""
                If Len(OldData11_1.Value) > 1400 Then strErrmsg &= "原計畫內容1資料過長有誤!" & vbCrLf
                If Len(TeacherName1.Text) > 1400 Then strErrmsg &= "原計畫內容2資料過長有誤!" & vbCrLf
                If Len(NewData11_1.Value) > 1400 Then strErrmsg &= "變更內容1資料過長有誤!" & vbCrLf
                If Len(TeacherName1_2.Text) > 1400 Then strErrmsg &= "變更內容2資料過長有誤!" & vbCrLf

                If strErrmsg <> "" Then
                    Try
                        '取得錯誤資訊寫入
                        strErrmsg &= "Path: TC_05_001_CHG" & vbCrLf
                        strErrmsg &= "FilePath: " & Request.FilePath & vbCrLf
                        strErrmsg &= "sql: " & sql & vbCrLf
                        strErrmsg &= "dt.Rows.Count: " & CStr(dt.Rows.Count) & vbCrLf
                        strErrmsg &= " Request(""PlanID""): " & rPlanID & vbCrLf
                        strErrmsg &= " Request(""cid""): " & rComIDNO & vbCrLf
                        strErrmsg &= " Request(""no""): " & rSeqNo & vbCrLf
                        strErrmsg &= "SubSeqNo: " & iSubSeqNO & vbCrLf
                        strErrmsg &= "CDate: " & CDate(ApplyDate.Text).ToString("yyyy/MM/dd") & vbCrLf
                        strErrmsg &= "AltDataID: " & v_ChgItem & vbCrLf '.SelectedValue 
                        strErrmsg &= "OldData11_1: " & OldData11_1.Value & vbCrLf
                        strErrmsg &= "TeacherName1: " & TeacherName1.Text & vbCrLf
                        strErrmsg &= "NewData11_1: " & NewData11_1.Value & vbCrLf
                        strErrmsg &= "TeacherName1_2: " & TeacherName1_2.Text & vbCrLf
                        strErrmsg &= "ReviseAcct: " & sm.UserInfo.UserID & vbCrLf
                        strErrmsg &= "ReviseCont: " & ReviseCont.Text & vbCrLf
                        strErrmsg &= "PlanKind: " & CStr(iPlanKind) & vbCrLf
                        'strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    Catch ex As Exception
                        Dim strErrmsg5 As String = ""
                        strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                        strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                        strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                        strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                        Call TIMS.WriteTraceLog(strErrmsg5)

                    End Try
                    Common.MessageBox(Page, strErrmsg)
                    Exit Sub
                End If


                If (ViewState(vs_UpdateTrainDesc) = "Y") Then
                    Dim htSEL As Hashtable = New Hashtable
                    TIMS.SetMyValue2(htSEL, "rPlanID", rPlanID) 'Request("PlanID")
                    TIMS.SetMyValue2(htSEL, "rComIDNO", rComIDNO) 'Request("cid")
                    TIMS.SetMyValue2(htSEL, "rSeqNo", rSeqNo) 'Request("no")
                    TIMS.SetMyValue2(htSEL, "SCDate", TIMS.Cdate3(ApplyDate.Text)) 'Request("CDate")
                    TIMS.SetMyValue2(htSEL, "SubSeqNo", iSubSeqNO) 'Request("SubNo")
                    htSEL.Add("RID", RIDValue.Value)
                    htSEL.Add("TECHIDs", NewData11_1.Value)
                    htSEL.Add("TechTYPE", "A")
                    CHK_REVISE_TEACHER(htSEL, strErrmsg)
                    If strErrmsg <> "" Then
                        Common.MessageBox(Page, strErrmsg)
                        Exit Sub
                    End If
                    SAVE_REVISE_TEACHER(htSEL, gobjconn)
                End If

                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i師資, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    dr("OldData11_1") = Left(OldData11_1.Value, 1400)
                    dr("OldData11_2") = Left(TeacherName1.Text, 1400)
                    dr("NewData11_1") = Left(NewData11_1.Value, 1400)
                    dr("NewData11_2") = Left(TeacherName1_2.Text, 1400)
                    dr("NewData11_3") = Hid_NewData11_3.Value
                    'htSS.Clear()
                    Dim htSS As Hashtable = New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)

                    '20080702 andy 課程表「師資」 start
                    '產學訓修改課程表師資前要先產生暫存的師資名單
                    ViewState(vs_UpdateItemIndex) = Cst_i師資
                    Select Case ViewState(vs_TEMP11_TrainDescDT)            'Session("TEMP11_TrainDescDT")
                        Case Is <> NewData11_1.Value     '暫存師資名單<>目前所選擇師資名單
                            ViewState(vs_TEMP11_TrainDescDT) = dt.Rows(0).Item("NewData11_1")
                        Case Is = NewData11_1.Value      '已產生暫存的師資名單and暫存師資名單=使用者選擇的師資名單
                            Select Case ViewState(vs_UpdateTrainDesc)
                                Case "Y" '已更改 Plan_TrainDesc
                                    DbAccess.UpdateDataTable(dt, da)
                                    ViewState(vs_TEMP11_TrainDescDT) = ""
                            End Select
                        Case Else
                            ViewState(vs_TEMP11_TrainDescDT) = dt.Rows(0).Item("NewData11_1") '產生暫存師資名單
                            ViewState(vs_UpdateTrainDesc) = "N" '尚未更改 Plan_TrainDesc
                    End Select

                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    strErrmsg &= ex.Message 'ex.ToString()
                    If flagDebugTest Then Throw ex
                End Try

            Case Cst_i助教 '產投儲存
                Dim strErrmsg As String = ""
                If Len(OldData20_1.Value) > 1400 Then strErrmsg &= "原計畫內容1資料過長有誤!" & vbCrLf
                If Len(TeacherName2.Text) > 1400 Then strErrmsg &= "原計畫內容2資料過長有誤!" & vbCrLf
                If Len(NewData20_1.Value) > 1400 Then strErrmsg &= "變更內容1資料過長有誤!" & vbCrLf
                If Len(TeacherName2_2.Text) > 1400 Then strErrmsg &= "變更內容2資料過長有誤!" & vbCrLf
                If strErrmsg <> "" Then
                    Common.MessageBox(Page, strErrmsg)
                    Exit Sub
                End If

                If (ViewState(vs_UpdateTrainDesc) = "Y") Then
                    Dim htSEL As Hashtable = New Hashtable
                    TIMS.SetMyValue2(htSEL, "rPlanID", rPlanID) 'Request("PlanID")
                    TIMS.SetMyValue2(htSEL, "rComIDNO", rComIDNO) 'Request("cid")
                    TIMS.SetMyValue2(htSEL, "rSeqNo", rSeqNo) 'Request("no")
                    TIMS.SetMyValue2(htSEL, "SCDate", TIMS.Cdate3(ApplyDate.Text)) 'Request("CDate")
                    TIMS.SetMyValue2(htSEL, "SubSeqNo", iSubSeqNO) 'Request("SubNo")
                    htSEL.Add("RID", RIDValue.Value)
                    htSEL.Add("TECHIDs", NewData20_1.Value)
                    htSEL.Add("TechTYPE", "B")
                    CHK_REVISE_TEACHER(htSEL, strErrmsg)
                    If strErrmsg <> "" Then
                        Common.MessageBox(Page, strErrmsg)
                        Exit Sub
                    End If
                    SAVE_REVISE_TEACHER(htSEL, gobjconn)
                End If

                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i助教, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    dr("OldData20_1") = Left(OldData20_1.Value, 1400)
                    dr("OldData20_2") = Left(TeacherName2.Text, 1400)
                    dr("NewData20_1") = Left(NewData20_1.Value, 1400)
                    dr("NewData20_2") = Left(TeacherName2_2.Text, 1400)
                    dr("NewData20_3") = Hid_NewData20_3.Value
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)

                    '課程表「助教」start
                    '產學訓修改課程表師資前要先產生暫存的師資名單
                    ViewState(vs_UpdateItemIndex) = Cst_i助教
                    Select Case ViewState(vs_TEMP20_TrainDescDT)
                        Case Is <> NewData20_1.Value     '暫存師資名單<>目前所選擇師資名單
                            ViewState(vs_TEMP20_TrainDescDT) = dt.Rows(0).Item("NewData20_1")
                        Case Is = NewData20_1.Value      '已產生暫存的師資名單and暫存師資名單=使用者選擇的師資名單
                            Select Case ViewState(vs_UpdateTrainDesc)
                                Case "Y" '已更改 PLAN_TRAINDESC
                                    DbAccess.UpdateDataTable(dt, da)
                                    ViewState(vs_TEMP20_TrainDescDT) = ""
                            End Select
                        Case Else
                            ViewState(vs_TEMP20_TrainDescDT) = dt.Rows(0).Item("NewData20_1") '產生暫存師資名單
                            ViewState(vs_UpdateTrainDesc) = "N" '尚未更改 Plan_TrainDesc
                    End Select

                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)
                    'strErrmsg += ex.ToString()
                    strErrmsg &= ex.Message
                    If flagDebugTest Then Throw ex
                End Try
                'If strErrmsg <> "" Then Common.MessageBox(Page, strErrmsg)

            Case Cst_i核定人數
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i核定人數, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData12_1") = OldData12_1.Text
                dr("NewData12_1") = NewData12_1.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

                If iPlanKind = 1 Then
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TNum") = NewData12_1.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TNum") = NewData12_1.Text
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i增班 '變更班數
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i增班, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData13_1") = OldData13_1.Text
                dr("NewData13_1") = NewData13_1.Text
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

                If iPlanKind = 1 Then
                    '2006/11/16 by Ellen update PlanInfo 
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    dt = DbAccess.GetDataTable(sql, da, gobjconn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("ClassCount") = NewData13_1.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i科場地 '變更學(術)科場地
                Dim v_NewData14_1b As String = TIMS.GetListValue(NewData14_1b) '學科場地地址1
                Dim v_NewData14_2b As String = TIMS.GetListValue(NewData14_2b) '術科場地地址1
                Dim v_NewData14_3 As String = TIMS.GetListValue(NewData14_3) '學科場地地址2
                Dim v_NewData14_4 As String = TIMS.GetListValue(NewData14_4) '術科場地地址2
                Dim drSciPc1 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_1b, 1, gobjconn) '取得學科場地的地址
                Dim drTechPc1 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_2b, 2, gobjconn) '取得術科場地的地址
                Dim drSciPc2 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_3, 1, gobjconn) '取得學科場地的地址2
                Dim drTechPc2 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_4, 2, gobjconn) '取得術科場地的地址2
                Hid_NewData8_4.Value = If((drSciPc1 IsNot Nothing), drSciPc1("PTID").ToString(), "") '學科場地地址
                Hid_NewData8_5.Value = If((drTechPc1 IsNot Nothing), drTechPc1("PTID").ToString(), "") '術科場地地址
                Hid_NewData8_6.Value = If((drSciPc2 IsNot Nothing), drSciPc2("PTID").ToString(), "") '學科場地地址2
                Hid_NewData8_7.Value = If((drTechPc2 IsNot Nothing), drTechPc2("PTID").ToString(), "") '術科場地地址2
                Dim strErrmsg As String = ""
                If v_NewData14_3 <> "" AndAlso v_NewData14_1b = "" Then strErrmsg &= "學科場地2 有選值，學科場地1(不可為空)!" & vbCrLf
                If v_NewData14_4 <> "" AndAlso v_NewData14_2b = "" Then strErrmsg &= "術科場地2 有選值，術科場地1(不可為空)!" & vbCrLf
                If v_NewData14_1b = "" AndAlso v_NewData14_2b = "" Then strErrmsg &= "學科場地1 或 術科場地1 (不可為空)!" & vbCrLf
                If v_NewData14_3 <> "" AndAlso v_NewData14_1b <> "" AndAlso v_NewData14_3 = v_NewData14_1b Then strErrmsg &= "學科場地1 與 學科場地2(不可為相同)!" & vbCrLf
                If v_NewData14_4 <> "" AndAlso v_NewData14_2b <> "" AndAlso v_NewData14_4 = v_NewData14_2b Then strErrmsg &= "術科場地1 與 術科場地2(不可為相同)!" & vbCrLf
                If strErrmsg <> "" Then
                    Common.MessageBox(Page, strErrmsg)
                    Exit Sub
                End If

                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i科場地, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData14_1") = If(OldData14_1b.Value <> "", OldData14_1b.Value, Convert.DBNull)
                dr("OldData14_2") = If(OldData14_2b.Value <> "", OldData14_2b.Value, Convert.DBNull)
                dr("OldData14_3") = If(OldData14_3.Value <> "", OldData14_3.Value, Convert.DBNull)
                dr("OldData14_4") = If(OldData14_4.Value <> "", OldData14_4.Value, Convert.DBNull)
                dr("OldData8_4") = If(Hid_OldData8_4.Value <> "", Hid_OldData8_4.Value, Convert.DBNull)
                dr("OldData8_5") = If(Hid_OldData8_5.Value <> "", Hid_OldData8_5.Value, Convert.DBNull)
                dr("OldData8_6") = If(Hid_OldData8_6.Value <> "", Hid_OldData8_6.Value, Convert.DBNull)
                dr("OldData8_7") = If(Hid_OldData8_7.Value <> "", Hid_OldData8_7.Value, Convert.DBNull)

                dr("NewData14_1") = If(v_NewData14_1b <> "", v_NewData14_1b, Convert.DBNull) '學科場地1
                dr("NewData14_2") = If(v_NewData14_2b <> "", v_NewData14_2b, Convert.DBNull) '術科場地1
                dr("NewData14_3") = If(v_NewData14_3 <> "", v_NewData14_3, Convert.DBNull) '學科場地2
                dr("NewData14_4") = If(v_NewData14_4 <> "", v_NewData14_4, Convert.DBNull) '術科場地2
                dr("NewData8_4") = If(Hid_NewData8_4.Value <> "", Hid_NewData8_4.Value, Convert.DBNull) '學科場地地址
                dr("NewData8_5") = If(Hid_NewData8_5.Value <> "", Hid_NewData8_5.Value, Convert.DBNull) '術科場地地址
                dr("NewData8_6") = If(Hid_NewData8_6.Value <> "", Hid_NewData8_6.Value, Convert.DBNull) '學科場地地址2
                dr("NewData8_7") = If(Hid_NewData8_7.Value <> "", Hid_NewData8_7.Value, Convert.DBNull) '術科場地地址2
                'dr("NewData8_4")=If(TaddressS2.SelectedIndex <= 0, Convert.DBNull, TaddressS2.SelectedValue)
                'dr("NewData8_5")=If(TaddressT2.SelectedIndex <= 0, Convert.DBNull, TaddressT2.SelectedValue)
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

                ViewState(vs_UpdateItemIndex) = Cst_i科場地 '20080715 andy

            Case Cst_i包班種類 '變更包班種類
                Select Case v_PackageTypeNew'PackageTypeNew.SelectedValue
                    Case "3"
                        If Session("Revise_BusPackage") Is Nothing OrElse DG_BusPackageNew.Items.Count = 0 Then
                            'Common.SetListItem(PackageTypeNew, "3")
                            Common.MessageBox(Page, "請輸入欲變更之包班種類!!")
                            Exit Sub
                        End If
                End Select

                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i包班種類, Cst_sql, gobjconn)
                'ByRef htSS As Hashtable 'Dim htSS As New Hashtable 'htSS Hashtable() 'htSS.Add("strSetId", strSetId)
                Dim htSS As New Hashtable From {
                    {"sql", sql},
                    {"rPlanID", rPlanID},
                    {"rComIDNO", rComIDNO},
                    {"rSeqNo", rSeqNo},
                    {"iSubSeqNO", iSubSeqNO},
                    {"sApplyDate", ApplyDate.Text},
                    {"sAltDataID", v_ChgItem}, 'ChgItem.SelectedValue)
                    {"iPlanKind", iPlanKind},
                    {"OldData4_1", hidPackageTypeOld.Value},
                    {"NewData4_1", v_PackageTypeNew}, 'PackageTypeNew.SelectedValue)
                    {"sReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}, 'changeReason.SelectedValue)
                    {"sPackageTypeNew", v_PackageTypeNew}, 'PackageTypeNew.SelectedValue)
                    {"txtUname", txtUname.Text},
                    {"txtIntaxno", txtIntaxno.Text},
                    {"txtUbno", txtUbno.Text},
                    {"rPARTREDUC1", rPARTREDUC1}
                }
                Using oConn As SqlConnection = DbAccess.GetConnection()
                    Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
                    Try
                        Call SAVE_REVISE_BUSPACKAGE(Me, htSS, oConn, oTrans)
                        DbAccess.CommitTrans(oTrans)
                    Catch ex As Exception
                        Dim strErrmsg5 As String = ""
                        strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                        strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                        strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                        strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                        Call TIMS.WriteTraceLog(strErrmsg5)

                        DbAccess.RollbackTrans(oTrans)
                        Call TIMS.CloseDbConn(oConn)
                        Common.MessageBox(Page, "計畫變更儲存失敗!-9B41")
                        If flagDebugTest Then Throw ex
                        Exit Sub
                    End Try
                    Call TIMS.CloseDbConn(oConn)
                End Using

            Case Cst_i上課時間 '變更上課時間
                If Session("REVISE_ONCLASS") Is Nothing Then
                    Common.MessageBox(Page, "變更內容-請輸入欲變更之上課時間!-DA08")
                    Exit Sub
                End If
                If DataGrid2.Items.Count = 0 Then
                    Common.MessageBox(Page, "變更內容-請輸入欲變更之上課時間!-D425")
                    Exit Sub
                End If

                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課時間, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da) 'DbAccess.UpdateDataTable(dt, da, Trans)

                Dim dtTemp As DataTable = Session("REVISE_ONCLASS")
                Dim htSEL As New Hashtable
                'htSS=New Hashtable
                TIMS.SetMyValue2(htSEL, "rPlanID", Convert.ToString(rPlanID)) '計畫PK
                TIMS.SetMyValue2(htSEL, "rComIDNO", rComIDNO) '計畫PK
                TIMS.SetMyValue2(htSEL, "rSeqNo", Convert.ToString(rSeqNo)) '計畫PK
                TIMS.SetMyValue2(htSEL, "SCDate", TIMS.Cdate3(ApplyDate.Text)) 'ApplyDate.Text
                TIMS.SetMyValue2(htSEL, "SubSeqNo", Convert.ToString(iSubSeqNO)) 'iSubSeqNO
                TIMS.SetMyValue2(htSEL, "ssUserID", Convert.ToString(sm.UserInfo.UserID))
                Call SAVE_REVISE_ONCLASS(htSEL, dtTemp, gobjconn)
                Call SAVE_REVISE_ONCLASS_OLD(htSEL, dtTemp, gobjconn, DataGrid1)

            Case Cst_i其他 '變更原計畫內容、變更內容
                Dim strErrmsg As String = ""
                OldData15_1.Text = TIMS.ClearSQM(OldData15_1.Text)
                NewData15_1.Text = TIMS.ClearSQM(NewData15_1.Text)
                'If OldData15_1.Text <> "" Then OldData15_1.Text=OldData15_1.Text.Trim
                'If NewData15_1.Text <> "" Then NewData15_1.Text=NewData15_1.Text.Trim
                If Len(OldData15_1.Text) > 1400 Then strErrmsg &= "原計畫內容資料過長(1400)有誤!" & vbCrLf
                If Len(NewData15_1.Text) > 1400 Then strErrmsg &= "變更內容資料過長(1400)有誤!" & vbCrLf
                If strErrmsg <> "" Then
                    Common.MessageBox(Page, strErrmsg)
                    Exit Sub
                End If

                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                dr("OldData15_1") = If(OldData15_1.Text = "", Convert.DBNull, OldData15_1.Text)
                dr("NewData15_1") = If(NewData15_1.Text = "", Convert.DBNull, NewData15_1.Text)
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

            Case Cst_i報名日期      '20080825 andy add 變更報名起訖
                Dim Time1 As String = New_SEnterDate.Text + " " + String.Format("{0:00}", CType(HR1.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM1.SelectedValue, Integer))
                Dim Time2 As String = New_FEnterDate.Text + " " + String.Format("{0:00}", CType(HR2.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM2.SelectedValue, Integer))
                '(檢查資料)
                If New_SEnterDate.Text = "" Or New_FEnterDate.Text = "" Then
                    Common.MessageBox(Page, "報名起訖日期欄位為空白！")
                    Exit Sub
                End If
                If CDate(Time1) > CDate(Time2) Then
                    Common.MessageBox(Page, "「起始時間」" + Time1 + "大於「結束時間」" + Time2 + "！")
                    Exit Sub
                ElseIf CDate(Time1) = CDate(Time2) Then
                    Common.MessageBox(Page, "「起始時間」" + Time1 + "不能與「結束時間」" + Time2 + "相同！")
                    Exit Sub
                End If

                '(存入申請異動項目)
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                '異動項目
                dr("OldData17_1") = If(Old_SEnterDate.Text = "", Convert.DBNull, Old_SEnterDate.Text)
                dr("NewData17_1") = If(Time1 = "", Convert.DBNull, Time1)
                dr("OldData17_2") = If(Old_FEnterDate.Text = "", Convert.DBNull, Old_FEnterDate.Text)
                dr("NewData17_2") = If(Time2 = "", Convert.DBNull, Time2)
                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)

                'Try
                '    DbAccess.CommitTrans(Trans)
                'Catch ex As Exception
                '    DbAccess.RollbackTrans(Trans)
                '    Call TIMS.CloseDbConn(conn)
                '    Common.MessageBox(Page, "計畫變更儲存失敗!-84D8")
                '    If flagDebugTest Then Throw ex
                '    Exit Sub
                'End Try

                If iPlanKind = 1 Then
                    Dim pms_1 As New Hashtable From {{"rPlanID", rPlanID}, {"rComIDNO", rComIDNO}, {"rSeqNo", rSeqNo}}
                    Dim sql_1c As String = ""
                    sql_1c &= " SELECT b.OCID"
                    sql_1c &= " FROM PLAN_PLANINFO a "
                    sql_1c &= " JOIN CLASS_CLASSINFO b ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNo=b.SeqNo "
                    sql_1c &= " WHERE b.IsSuccess='Y'"
                    sql_1c &= " AND a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNo=@SeqNo"
                    Dim dtC As DataTable = DbAccess.GetDataTable(sql_1c, gobjconn, pms_1)
                    If dtC Is Nothing OrElse dtC.Rows.Count = 0 Then
                        Common.MessageBox(Page, "查無轉班後資料，儲存失敗！")
                        Exit Sub
                    End If
                    Dim drC1 As DataRow = dtC.Rows(0)
                    Dim u_OCID As String = Convert.ToString(drC1("OCID"))

                    If New_SEnterDate.Text <> "" OrElse New_FEnterDate.Text <> "" Then
                        Dim pms_u1 As New Hashtable From {{"ModifyAcct", sm.UserInfo.UserID}}
                        If New_SEnterDate.Text <> "" Then pms_u1.Add("SEnterDate", Time1)
                        If New_FEnterDate.Text <> "" Then pms_u1.Add("FEnterDate", Time2)
                        'M: 修改(最後異動狀態)
                        Dim u_sql As String = ""
                        u_sql &= " UPDATE CLASS_CLASSINFO"
                        u_sql &= " SET LastState='M',ModifyAcct=@ModifyAcct,ModifyDate=GETDATE()"
                        If New_SEnterDate.Text <> "" Then u_sql &= " ,SEnterDate=@SEnterDate"
                        If New_FEnterDate.Text <> "" Then u_sql &= " ,FEnterDate=@FEnterDate"
                        u_sql &= " WHERE OCID=@OCID"
                        DbAccess.ExecuteNonQuery(u_sql, gobjconn, pms_u1)
                    End If

                End If

            Case Cst_i遠距教學
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i遠距教學, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                Hid_DISTANCE.Value = TIMS.ClearSQM(Hid_DISTANCE.Value)
                Dim vrbl_DISTANCE As String = TIMS.GetListValue(rbl_DISTANCE)
                dr("OldData22_1") = If(Hid_DISTANCE.Value = "", Convert.DBNull, Hid_DISTANCE.Value)
                dr("NewData22_1") = If(vrbl_DISTANCE = "", Convert.DBNull, vrbl_DISTANCE)

                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)
                ViewState(vs_UpdateItemIndex) = Cst_i遠距教學

            Case Cst_i課程表  '20080715 andy
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i課程表, Cst_sql, gobjconn)
                dt = DbAccess.GetDataTable(sql, da, gobjconn)
                Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                Dim htSS As New Hashtable From {
                    {"ReviseCont", ReviseCont.Text},
                    {"changeReason", v_changeReason}
                }
                Call INS_CMN1(dr, sm, iPlanKind, htSS)
                DbAccess.UpdateDataTable(dt, da)
                '20070715 andy
                ViewState(vs_UpdateItemIndex) = Cst_i課程表
        End Select

        'Session("_search")=ViewState("_search")
        Dim strMassage2 As String = ""
        Select Case TIMS.CINT1(v_ChgItem)'ChgItem.SelectedValue
            Case Cst_i訓練期間, Cst_i課程表
                '2024'確認單位所選「課程大綱日期」與「上課時間之星期」是否正確，於基本儲存、正式儲存時
                Dim iPTDRID As Integer = Get_PTDRID(2, rPlanID, rComIDNO, rSeqNo, Get_NowDate(), iSubSeqNO, v_ChgItem, gobjconn)
                If iPTDRID > 0 Then
                    Dim dtDesc As DataTable = Get_PlanTrainDescNewRevise(iPTDRID, v_ChgItem) '課程表申請變更後
                    Dim dtOnClass As DataTable = GET_PLAN_ONCLASS_S1(gobjconn)
                    Dim hMsg1 As New Hashtable
                    If Not TIMS.CHK_WEEKDAY1(dtDesc, dtOnClass, hMsg1) Then
                        Dim s_STRAIN As String = TIMS.GetMyValue2(hMsg1, "STRAINDATE")
                        Dim s_WEEKS As String = TIMS.GetMyValue2(hMsg1, "WEEKS")
                        '所選的「課程日期」(2025/01/16)(星期四)與原「上課時段」之星期不符合，請再變更「上課時間」
                        strMassage2 &= $"所選的「課程日期」({s_STRAIN})({s_WEEKS})與原「上課時段」之星期不符合，請再變更「上課時間或課程表」!\n"
                    End If
                End If

            Case Cst_i上課時間 'REVISE_ONCLASS
                '2024'確認單位所選「課程大綱日期」與「上課時間之星期」是否正確，於基本儲存、正式儲存時
                Dim dtDesc As DataTable = GET_PLAN_TRAINDESC_S1(gobjconn)
                Dim dtOnClass As DataTable = GET_REVISE_ONCLASS_S1(gobjconn)
                Dim hMsg1 As New Hashtable
                If Not TIMS.CHK_WEEKDAY1(dtDesc, dtOnClass, hMsg1) Then
                    Dim s_STRAIN As String = TIMS.GetMyValue2(hMsg1, "STRAINDATE")
                    Dim s_WEEKS As String = TIMS.GetMyValue2(hMsg1, "WEEKS")
                    '所選的「課程日期」(2025/01/16)(星期四)與原「上課時段」之星期不符合，請再變更「上課時間」
                    strMassage2 &= $"所選的「課程日期」({s_STRAIN})({s_WEEKS})與原「上課時段」之星期不符合，請再變更「上課時間或課程表」!\n"
                End If

        End Select

        Dim okmsg As String = "計劃變更申請成功!\n"
        If strMassage <> "" Then okmsg = strMassage & "計劃變更申請成功!\n"
        If strMassage2 <> "" Then okmsg &= strMassage2
        'Call TIMS.CloseDbConn(conn)

        Dim rtn As String = ""
        Dim blFlag As Boolean = False
        Select Case v_ChgItem'ChgItem.SelectedValue
            Case Cst_i師資, Cst_i助教
                If ViewState(vs_UpdateTrainDesc) = "Y" Then
                    ViewState(vs_UpdateTrainDesc) = ""
                    blFlag = True
                End If
            Case Else
                ViewState(vs_UpdateTrainDesc) = ""
                blFlag = True
        End Select

        If blFlag Then
            '直接回上頁
            rtn = String.Concat("<script>", "blockAlert('", okmsg.Replace("\\n", "<br>"), "','',function(){ ")
            rtn += String.Concat("location.href='TC/05/TC_05_001.aspx?ID=", TIMS.ClearSQM(Request("ID")), "';} ", ");", "</script>")
            Page.RegisterStartupScript("msg", rtn)
        End If
    End Sub

    Function GET_PLAN_TRAINDESC_S1(oConn As SqlConnection) As DataTable
        Dim PMS1 As New Hashtable From {
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNO", rSeqNo}
        }
        Dim RSql As String = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
        Dim dtS1 As DataTable = DbAccess.GetDataTable(RSql, oConn, PMS1)
        Return dtS1
    End Function
    Function GET_REVISE_ONCLASS_S1(oConn As SqlConnection) As DataTable
        Dim V_SCDate As String = Get_NowDate()
        Dim PMS1 As New Hashtable From {
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNO", rSeqNo},
            {"SCDate", TIMS.Cdate2(V_SCDate)},
            {"SubSeqNO", iSubSeqNO}
        }
        Dim RSql As String = ""
        RSql &= " SELECT * FROM REVISE_ONCLASS "
        RSql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO AND SCDate=@SCDate AND SubSeqNO=@SubSeqNO"
        Dim dtS1 As DataTable = DbAccess.GetDataTable(RSql, oConn, PMS1)
        Return dtS1
    End Function

    Function GET_PLAN_ONCLASS_S1(oConn As SqlConnection) As DataTable
        Dim PMS1 As New Hashtable From {
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNO", rSeqNo}
        }
        Dim RSql As String = " SELECT * FROM PLAN_ONCLASS WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
        Dim dtS1 As DataTable = DbAccess.GetDataTable(RSql, oConn, PMS1)
        Return dtS1
    End Function

    '檢查是否存在尚在審核中的節次 
    Private Function ChkIsClassMatch(ByVal ClassID As String, ByVal SchoolDate As String, ByRef oConn As SqlConnection, ByRef oTrans As SqlTransaction) As String  'ClassID:節次
        Dim rst As String = ""
        Try
            Dim da As New SqlDataAdapter
            Dim dt As New DataTable
            Dim dr As DataRow
            Dim sql As String = ""
            sql = ""
            sql &= " SELECT * FROM PLAN_REVISE" & vbCrLf
            sql &= " WHERE AltDataID=2" & vbCrLf
            sql &= " AND ReviseStatus IS NULL" & vbCrLf
            sql &= " AND (1!=1" & vbCrLf
            sql &= "     OR OldData2_1=" & TIMS.To_date(SchoolDate) & vbCrLf 'CDate(SchoolDate).ToString("yyyy/MM/dd") 
            sql &= "     OR NewData2_1=" & TIMS.To_date(SchoolDate) & vbCrLf 'convert(varchar,NewData2_1,111)='" & CDate(SchoolDate).ToString("yyyy/MM/dd") & "'" & vbCrLf
            sql &= "    )" & vbCrLf
            'sql &= "   AND (convert(varchar,OldData2_1,111)='" & Convert.ToDateTime(SchoolDate).ToString("yyyy/MM/dd") & "'  OR  convert(varchar( 10),NewData2_1,111)='" & Convert.ToDateTime(SchoolDate).ToString("yyyy/MM/dd") & "' )" & vbCrLf
            sql &= " AND PlanID= " & rPlanID 'Request("PlanID") & vbCrLf
            sql &= " AND ComIDNO='" & rComIDNO & "'" 'Request("cid") & vbCrLf
            sql &= " AND SeqNO=" & rSeqNo 'Request("no") & vbCrLf
            With da
                .SelectCommand = New SqlCommand(sql, oConn, oTrans)
                .Fill(dt)
            End With
            If dt.Rows.Count > 0 Then
                For i As Int16 = 0 To dt.Rows.Count - 1
                    dr = dt.Rows(i)
                    Dim ClassAry_N As Array
                    Dim ClassAry_O As Array
                    Dim ClassAry_Chk As Array
                    Dim IsExist As Boolean = False
                    ClassAry_Chk = Split(ClassID, ",")
                    If Convert.ToDateTime(dr("NewData2_1")).ToString("yyyy/MM/dd") = Convert.ToDateTime(SchoolDate).ToString("yyyy/MM/dd") Then
                        ClassAry_N = Split(Convert.ToString(dr("NewData2_3")), ",")
                        If ClassAry_N.Length > 0 Then
                            For Each class1 As String In ClassAry_N
                                IsExist = False
                                For Each class2 As String In ClassAry_Chk
                                    If class1 = class2 Then IsExist = True '待審核中的節次與申請的節次一樣
                                Next
                                If IsExist = True Then
                                    rst = Convert.ToDateTime(Convert.ToString(dr("CDATE"))).ToString("yyyy/MM/dd")
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    If Convert.ToDateTime(dr("OldData2_1")).ToString("yyyy/MM/dd") = Convert.ToDateTime(SchoolDate).ToString("yyyy/MM/dd") Then
                        ClassAry_O = Split(Convert.ToString(dr("OldData2_3")), ",")
                        If ClassAry_O.Length > 0 Then
                            For Each class1 As String In ClassAry_O
                                IsExist = False
                                For Each class2 As String In ClassAry_Chk
                                    If class1 = class2 Then IsExist = True '待審核中的節次與申請的節次一樣
                                Next
                                If IsExist = True Then
                                    rst = Convert.ToDateTime(Convert.ToString(dr("CDATE"))).ToString("yyyy-MM-dd") '申請中有相同課程存在的日期
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Next
            End If
            'Call TIMS.CloseDbConn(conn)
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Common.MessageBox(Page, "發生錯誤:" & ex.Message.ToString())
        End Try
        Return rst
    End Function

    '(請記得於本變更申請審核通過後，修正課程表。)
    Function Get_MaxMinDate(ByVal OCID1 As String, ByVal SchoolDate1 As String, ByVal SchoolDate2 As String) As DataTable
        Dim rst As DataTable = Nothing
        Dim sql As String = ""
        sql &= " SELECT MinDate, MaxDate" & vbCrLf
        sql &= " FROM (SELECT MIN(SchoolDate) MinDate ,MAX(SchoolDate) MaxDate" & vbCrLf
        sql &= "    FROM CLASS_SCHEDULE" & vbCrLf
        sql &= "    WHERE 1=1" & vbCrLf
        If SchoolDate1 <> "" And SchoolDate2 <> "" And IsDate(SchoolDate1) And IsDate(SchoolDate2) Then
            'sql &= " AND (SchoolDate < '" & SchoolDate1 & "' OR SchoolDate > '" & SchoolDate2 & "' )" & vbCrLf
            sql &= "  AND (SchoolDate < " & TIMS.To_date(SchoolDate1) & " OR SchoolDate > " & TIMS.To_date(SchoolDate2) & " )" & vbCrLf
        End If
        sql &= "      AND Formal='Y' AND OCID='" & OCID1 & "'" & vbCrLf
        sql &= "      AND (class1 IS NOT NULL OR class2 IS NOT NULL OR class3 IS NOT NULL OR class4 IS NOT NULL" & vbCrLf
        sql &= "           OR class5 IS NOT NULL OR class6 IS NOT NULL OR class7 IS NOT NULL OR class8 IS NOT NULL" & vbCrLf
        sql &= "           OR class9 IS NOT NULL OR class10 IS NOT NULL OR class11 IS NOT NULL OR class12 IS NOT NULL)" & vbCrLf
        sql &= " ) G" & vbCrLf
        sql &= " WHERE MinDate IS NOT NULL AND MaxDate IS NOT NULL" & vbCrLf
        rst = DbAccess.GetDataTable(sql, gobjconn)
        Return rst
    End Function

    'TIMS 處存檢核
    Function CheckSaveData1(ByRef sErrMsg As String) As Boolean
        Dim rst As Boolean = True
        sErrMsg = ""

        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Select Case v_ChgItem'ChgItem.SelectedValue
            Case Cst_i訓練期間 '1
                If ASDate.Text = "" Or AEDate.Text = "" Then
                    sErrMsg &= "請輸入變更內容的起迄日期!"
                    Return False
                End If
                Dim vHR3 As String = ""
                Dim vMM3 As String = ""
                Call TIMS.sUtl_GetHRMM(vHR3, vMM3, HR3.SelectedValue, MM3.SelectedValue)
                Dim vHR4 As String = ""
                Dim vMM4 As String = ""
                Call TIMS.sUtl_GetHRMM(vHR4, vMM4, HR4.SelectedValue, MM4.SelectedValue)
                Dim Time1 As String = "" 'New_SEnterDate2.Text & " " & vHR3 & ":" & vMM3 'yyyy/MM/dd hh24:mi
                Dim Time2 As String = "" 'New_FEnterDate2.Text & " " & vHR4 & ":" & vMM4 'yyyy/MM/dd hh24:mi
                If New_SEnterDate2.Text <> "" Then Time1 = New_SEnterDate2.Text & " " & vHR3 & ":" & vMM3 'yyyy/MM/dd hh24:mi
                If New_SEnterDate2.Text <> "" Then Time2 = New_FEnterDate2.Text & " " & vHR4 & ":" & vMM4 'yyyy/MM/dd hh24:mi
                'Dim Time1, Time2 As String
                'Dim Time1 As String=New_SEnterDate2.Text + " " + String.Format("{0:00}", CType(HR3.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM3.SelectedValue, Integer))
                'Dim Time2 As String=New_FEnterDate2.Text + " " + String.Format("{0:00}", CType(HR4.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM4.SelectedValue, Integer))

                '非產投計畫(TIMS)
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                    If New_SEnterDate2.Text = "" Then
                        sErrMsg &= "「報名開始日期」為必填欄位!"
                        Return False
                    End If
                    If New_FEnterDate2.Text = "" Then
                        sErrMsg &= "「報名結束日期」為必填欄位!"
                        Return False
                    End If
                    'If New_SEnterDate2.Text="" Or New_FEnterDate2.Text="" Then
                    '    sErrMsg &= "報名起訖日期欄位為空白!"
                    '    Return False
                    'End If
                    If Time1 <> "" AndAlso Time2 <> "" AndAlso CDate(Time1) >= CDate(Time2) Then
                        sErrMsg &= "「報名起始時間」" + Time1 + " 大於或等於「報名結束時間」" + Time2
                        Return False
                    ElseIf Time1 <> "" AndAlso ASDate.Text <> "" AndAlso CDate(Time1) >= CDate(ASDate.Text) Then
                        sErrMsg &= "「報名起始時間」" + Time1 + " 大於或等於「開訓時間」" + (CDate(ASDate.Text)).ToString("yyyy/MM/dd HH:mm@ss")
                        Return False
                    ElseIf Time2 <> "" AndAlso ASDate.Text <> "" AndAlso CDate(Time2) >= CDate(ASDate.Text) Then
                        sErrMsg &= "「報名結束時間」" + Time2 + " 大於或等於「開訓時間」" + (CDate(ASDate.Text)).ToString("yyyy/MM/dd HH:mm@ss")
                        Return False
                    End If

                    'New_ExamPeriod.SelectedValue
                    Dim v_New_ExamPeriod As String = TIMS.GetListValue(New_ExamPeriod)
                    If Not flag_TPlanID70_1 Then
                        If New_Examdate.Text = "" Then
                            sErrMsg &= "班別資料「甄試日期」為必填欄位!" & vbCrLf
                            Return False
                        ElseIf New_Examdate.Text <> "" AndAlso v_New_ExamPeriod = "" Then
                            sErrMsg &= "班別資料「甄試日期」 時段：全天、上午、下午 時段請擇一選擇!" & vbCrLf '20100329 add 甄試時段
                            Return False
                        End If
                    End If

                    If New_CheckInDate.Text = "" Then
                        sErrMsg &= "「報到日期」為必填欄位!" & vbCrLf
                        Return False
                    End If

                    '首頁>>訓練機構管理>>班級變更申請 'TC/05/TC_05_001
                    '變更項目-訓練期間，原計畫內容與變更內容內新增「甄試日期」，為必填。
                    '「報名開始日期」需早於「報名結束日期」，「報名結束日期」需早於「甄試日期」2天，例如:報名截止日：104/09/09，最快可辦理甄試的日期：104/09/11，「甄試日期」需早於「報到日期」，「報到日期」最晚可與開訓日同一天，另需顯示"報名登錄最晚可作業日期：104/09/09"。 
                    '變更項目訓練期間內的「甄試日期」最快得安排於報名截止當日起2日後。(ex:報名截止日：104/09/09，最快可辦理甄試的日期：104/09/11)
                    '(ex：1.報名截止日：104/09/09，甄試日：104/09/11，則報名登錄最晚可作業日期：104/09/09
                    '     2.報名截止日：104/09/09，甄試日：104/09/15，則報名登錄最晚可作業日期：104/09/12)。
                    If New_Examdate.Text <> "" AndAlso New_FEnterDate2.Text <> "" AndAlso (CDate(New_Examdate.Text) <= CDate(New_FEnterDate2.Text)) Then
                        sErrMsg &= "班別資料「甄試日期」必須大於「報名結束日期」!"
                        Return False
                    End If
                    If New_Examdate.Text <> "" AndAlso New_FEnterDate2.Text <> "" AndAlso DateDiff(DateInterval.Day, CDate(New_FEnterDate2.Text), CDate(New_Examdate.Text)) < 2 Then
                        '「甄試日期」最快得安排於報名截止當日起2日後。
                        sErrMsg &= "班別資料「甄試日期」最快得安排於「報名結束日期」當日起2日後!"
                        Return False
                    End If

                    If RIDValue.Value <> "" Then Hid_RID1.Value = Convert.ToString(RIDValue.Value).Substring(0, 1)
                    If Hid_RID1.Value = "" Then Hid_RID1.Value = Convert.ToString(sm.UserInfo.RID).Substring(0, 1)
                    If New_Examdate.Text <> "" AndAlso TIMS.Chk_HOLDATE(Hid_RID1.Value, New_Examdate.Text, gobjconn) Then
                        sErrMsg &= "班別資料「甄試日期」不可為例假日!" & vbCrLf
                        Return False
                    ElseIf New_CheckInDate.Text <> "" AndAlso TIMS.Chk_HOLDATE(Hid_RID1.Value, New_CheckInDate.Text, gobjconn) Then
                        sErrMsg &= "班別資料「報到日期」不可為例假日!" & vbCrLf
                        Return False
                    ElseIf New_Examdate.Text <> "" AndAlso ASDate.Text <> "" AndAlso CDate(New_Examdate.Text) >= CDate(ASDate.Text) Then
                        sErrMsg &= String.Concat("「甄試日期」", TIMS.Cdate3(New_Examdate.Text), " 大於或等於「開訓時間」", TIMS.Cdate3(ASDate.Text)) '(CDate(ASDate.Text)).ToString("yyyy/MM/dd HH:mm@ss")
                        Return False
                    ElseIf New_CheckInDate.Text <> "" AndAlso ASDate.Text <> "" AndAlso CDate(New_CheckInDate.Text) > CDate(ASDate.Text) Then
                        sErrMsg &= String.Concat("「報到日期」", TIMS.Cdate3(New_CheckInDate.Text), " 大於「開訓時間」", TIMS.Cdate3(ASDate.Text)) '().ToString("yyyy/MM/dd HH:mm@ss")
                        Return False
                    ElseIf New_Examdate.Text <> "" AndAlso New_CheckInDate.Text <> "" AndAlso CDate(New_Examdate.Text) >= CDate(New_CheckInDate.Text) Then
                        sErrMsg &= String.Concat("「甄試日期」", TIMS.Cdate3(New_Examdate.Text), " 大於或等於「報到日期」", TIMS.Cdate3(New_CheckInDate.Text))
                        Return False
                    End If

                    '(所有日期不可為空) All dates cannot be empty
                    Dim fg_ALLDATENOTEMPTY1 As Boolean = (ASDate.Text <> "" AndAlso BSDate.Text <> "" AndAlso AEDate.Text <> "" AndAlso BEDate.Text <> "" _
                        AndAlso Old_Examdate.Text <> "" AndAlso New_Examdate.Text <> "" AndAlso Old_CheckInDate.Text <> "" AndAlso New_CheckInDate.Text <> "" _
                        AndAlso Old_SEnterDate2.Text <> "" AndAlso Time1 <> "" AndAlso Old_FEnterDate2.Text <> "" AndAlso Time2 <> "")
                    If fg_ALLDATENOTEMPTY1 Then
                        If CDate(ASDate.Text) = CDate(BSDate.Text) AndAlso CDate(AEDate.Text) = CDate(BEDate.Text) _
                        AndAlso CDate(Old_Examdate.Text) = CDate(New_Examdate.Text) AndAlso CDate(Old_CheckInDate.Text) = CDate(New_CheckInDate.Text) _
                        AndAlso CDate(Old_SEnterDate2.Text) = CDate(Time1) AndAlso CDate(Old_FEnterDate2.Text) = CDate(Time2) Then
                            sErrMsg &= String.Concat(" 「訓練起迄日」", TIMS.Cdate3(ASDate.Text), "~", TIMS.Cdate3(AEDate.Text), "或「報名起迄日」", TIMS.Cdate3(Old_SEnterDate2.Text), "~", TIMS.Cdate3(Old_FEnterDate2.Text), " 不能與舊日期的相同!")
                            Return False
                        End If
                    End If
                Else
                    '產投(非TIMS)
                    If CDate(ASDate.Text) = CDate(BSDate.Text) AndAlso CDate(AEDate.Text) = CDate(BEDate.Text) Then
                        sErrMsg &= String.Concat("新的「訓練起迄日」", TIMS.Cdate3(ASDate.Text), "~", TIMS.Cdate3(AEDate.Text), "不能與舊日期的相同!")
                        Return False
                    End If
                End If
        End Select

        If sErrMsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>TIMS儲存-非產投</summary>
    Sub Save_Sub()
        Dim sErrMsg As String = ""
        Call CheckSaveData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        'New_ExamPeriod.SelectedValue
        Dim v_New_ExamPeriod As String = TIMS.GetListValue(New_ExamPeriod)
        Dim v_changeReason As String = TIMS.GetListValue(changeReason)
        Dim chkstr As String = ""
        Dim TotalItem As String = ""
        Dim campare As String = ""

        Dim strMassage As String = "" '檢查是否為全日制的錯誤訊息
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = ""
        sql = " SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        Dim iPlanKind As Integer = DbAccess.ExecuteScalar(sql, gobjconn)

        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem) 'ChgItem.SelectedValue 
        If v_ChgItem = "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        iSubSeqNO = GET_MaxSUBSEQNO_28(v_ChgItem)
        If (ReviseCont.Text.Length > 255) Then ReviseCont.Text = TIMS.Get_Substr1(ReviseCont.Text, 255)

        ''20090320 andy edit start
        'chkstr=""
        'chkstr &= " SELECT MAX(SubSeqNo) MAXSubSeqNo FROM PLAN_REVISE "
        'chkstr &= " WHERE 1=1 "
        'chkstr &= " AND PlanID=" & TIMS.ClearSQM(rPlanID)
        'chkstr &= " AND ComIDNO='" & TIMS.ClearSQM(rComIDNO) & "'"
        'chkstr &= " AND SeqNO=" & TIMS.ClearSQM(rSeqNo)
        'chkstr &= " AND CDate=" & TIMS.to_date(ApplyDate.Text)
        'Dim objValue As Object=DbAccess.ExecuteScalar(chkstr, gobjconn)
        ''Dim iSubSeqNO As Integer=1 'Dim iSubSeqNO As Integer=0
        'If objValue IsNot Nothing Then
        '    iSubSeqNO=If(Convert.ToString(objValue) <> "", (Val(objValue) + 1), 1) '有值+1,空的初始值為1
        'End If
        ''(查無所有類案序號預設為1)
        'If (iSubSeqNO=0) Then iSubSeqNO=1 '空的初始值為1

        Dim ck_OCID As String = "" '已轉班
        Dim drPP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, gobjconn)
        If drPP IsNot Nothing Then ck_OCID = Convert.ToString(drPP("OCID"))

        Dim TransConn As SqlConnection = DbAccess.GetConnection()
        'Call TIMS.OpenDbConn(TransConn)

        Select Case TIMS.CINT1(v_ChgItem)
            Case Cst_i訓練期間 'As Integer=1
                ' start  (檢查資料)
                Dim vHR3 As String = ""
                Dim vMM3 As String = ""
                Call TIMS.sUtl_GetHRMM(vHR3, vMM3, HR3.SelectedValue, MM3.SelectedValue)
                Dim vHR4 As String = ""
                Dim vMM4 As String = ""
                Call TIMS.sUtl_GetHRMM(vHR4, vMM4, HR4.SelectedValue, MM4.SelectedValue)
                'Dim Time1 As String=New_SEnterDate2.Text & " " & vHR3 & ":" & vMM3
                'Dim Time2 As String=New_FEnterDate2.Text & " " & vHR4 & ":" & vMM4 ' String.Format("{0:00}", CType(HR4.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM4.SelectedValue, Integer))
                'yyyy/MM/dd hh24:mi
                Dim Time1 As String = "" 'New_SEnterDate2.Text & " " & vHR3 & ":" & vMM3 'yyyy/MM/dd hh24:mi
                Dim Time2 As String = "" 'New_FEnterDate2.Text & " " & vHR4 & ":" & vMM4 'yyyy/MM/dd hh24:mi
                If New_SEnterDate2.Text <> "" Then Time1 = $"{New_SEnterDate2.Text} {vHR3}:{vMM3}" 'yyyy/MM/dd hh24:mi
                If New_SEnterDate2.Text <> "" Then Time2 = $"{New_FEnterDate2.Text} {vHR4}:{vMM4}" 'yyyy/MM/dd hh24:mi

                '20080804 andy edit 
                'Dim ck_OCID As String=""
                'Dim drP As DataRow=TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, gobjconn)
                'If Not drP Is Nothing Then ck_OCID=Convert.ToString(drP("ocid"))
                '(請記得於本變更申請審核通過後，修正課程表。)
                If ck_OCID <> "" Then
                    Dim dt2M As DataTable = Get_MaxMinDate(ck_OCID, ASDate.Text, AEDate.Text)
                    If dt2M.Rows.Count > 0 Then strMassage = Cst_msg1
                End If

                '(存入異動申請項目)
                sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練期間, Cst_sql, gobjconn)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目 (開結訓日期)
                    dr("OldData1_1") = BSDate.Text
                    dr("OldData1_2") = BEDate.Text
                    dr("NewData1_1") = ASDate.Text
                    dr("NewData1_2") = AEDate.Text
                    '20081107 異動項目(報名起迄)
                    dr("OldData17_1") = If(Old_SEnterDate2.Text = "", Convert.DBNull, Old_SEnterDate2.Text)
                    dr("NewData17_1") = If(Time1 = "", Convert.DBNull, Time1) 'yyyy/MM/dd hh24:mi
                    dr("OldData17_2") = If(Old_FEnterDate2.Text = "", Convert.DBNull, Old_FEnterDate2.Text)
                    dr("NewData17_2") = If(Time2 = "", Convert.DBNull, Time2) 'yyyy/MM/dd hh24:mi
                    '20160129 
                    dr("OLDDATA3_1") = If(Old_Examdate.Text = "", Convert.DBNull, Old_Examdate.Text)
                    dr("NEWDATA3_1") = If(New_Examdate.Text = "", Convert.DBNull, New_Examdate.Text)
                    dr("OLDDATA10_1") = If(HidOld_ExamPeriod.Value = "", Convert.DBNull, HidOld_ExamPeriod.Value)
                    dr("NEWDATA10_1") = If(v_New_ExamPeriod = "", Convert.DBNull, v_New_ExamPeriod)
                    dr("OLDDATA2_1") = If(Old_CheckInDate.Text = "", Convert.DBNull, Old_CheckInDate.Text)
                    dr("NEWDATA2_1") = If(New_CheckInDate.Text = "", Convert.DBNull, New_CheckInDate.Text)
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    If iPlanKind = 1 Then
                        sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("STDate") = ASDate.Text
                            dr("FDDate") = AEDate.Text
                            If Time1 <> "" Then dr("SEnterDate") = Time1
                            If Time2 <> "" Then dr("FEnterDate") = Time2
                            dr("Examdate") = If(New_Examdate.Text <> "", New_Examdate.Text, Convert.DBNull) '(必填)
                            dr("ExamPeriod") = v_New_ExamPeriod 'New_ExamPeriod.SelectedValue '(必填)
                            dr("CheckInDate") = New_CheckInDate.Text '(必填)
                            'If New_Examdate.Text <> "" Then dr("Examdate")=New_Examdate.Text
                            'If New_ExamPeriod.SelectedValue <> "" Then dr("ExamPeriod")=New_ExamPeriod.SelectedValue
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da, Trans)
                            'DbAccess.CommitTrans(Trans)
                        End If

                        sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                        '2006/03/ add conn by matt
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0) '報名起迄
                            If Time1 <> "" Then dr("SEnterDate") = Time1
                            If Time2 <> "" Then dr("FEnterDate") = Time2
                            dr("Examdate") = If(New_Examdate.Text <> "", New_Examdate.Text, Convert.DBNull) '(必填)
                            dr("ExamPeriod") = v_New_ExamPeriod '.SelectedValue '(必填)
                            dr("CheckInDate") = New_CheckInDate.Text '(必填)
                            If Time2 <> "" AndAlso New_Examdate.Text <> "" Then
                                'Hid_RID1.Value=dr("RID")
                                Hid_RID1.Value = Convert.ToString(dr("RID")).Substring(0, 1)
                                Dim sFENTERDATE As String = Time2
                                Dim sEXAMDATE As String = New_Examdate.Text
                                Dim SS1 As String = ""
                                TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
                                Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, gobjconn)
                                'Dim sFENTERDATE2 As String=TIMS.GET_FENTERDATE2(Hid_RID1.Value, sFENTERDATE, sEXAMDATE, gobjconn)
                                If sFENTERDATE2 <> "" Then dr("FEnterDate2") = sFENTERDATE2 'TIMS.GET_FENTERDATE2(Time2, New_Examdate.Text)
                            End If
                            dr("STDate") = ASDate.Text
                            dr("FTDate") = AEDate.Text
                            dr("LastState") = "M" 'M: 修改(最後異動狀態)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da, Trans)
                            'DbAccess.CommitTrans(Trans)
                        End If

                        'Dim i As Integer
                        sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & ViewState(vs_OCID) & "'"
                        '2006/03/ add conn by matt
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        If dt.Rows.Count <> 0 Then
                            For i As Integer = 0 To dt.Rows.Count - 1
                                dr = dt.Rows(i)
                                dr("OpenDate") = ASDate.Text
                                dr("CloseDate") = AEDate.Text
                                dr("ModifyAcct") = sm.UserInfo.UserID
                                dr("ModifyDate") = Now()
                                DbAccess.UpdateDataTable(dt, da, Trans)
                            Next
                        End If

                        '更新課程資料
                        'Dim myThreadDelegate As New Threading.ThreadStart(AddressOf SetClassSchedule)
                        'Dim myThread As New Threading.Thread(myThreadDelegate)
                        'myThread.Start()
                        sql = "UPDATE STUD_DATALID SET RESULTDATE=" & TIMS.To_date(AEDate.Text) & " where OCID='" & ViewState(vs_OCID) & "'"
                        DbAccess.ExecuteNonQuery(sql, Trans)
                    End If
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-3D77")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i訓練時段 'As Integer=2
                '訓練時段
                '20080912  Andy edit
                '20080912  Andy edit
                '(檢查資料)
                'Dim Sqlda As New SqlDataAdapter
                'Dim sqlconn As New SqlConnection
                'Dim i As Integer=0
                'Dim n As Integer=0
                'Dim x As Integer=0
                'Dim str1, str2 As String
                'Dim Selstr1, Selstr2, ClassID_1, ClassID_2, TeachID_1, TeachID_2 As String
                'Dim SelClassStr, SelClassStr2, SelTeachStr, SelTeachStr2 As String
                'Dim TargetLB3 As New ListBox
                ''Dim sqlcmd As New SqlCommand
                'Dim objTrans As SqlTransaction=Nothing

                Dim Selstr1 As String = ""
                Dim ClassID_1 As String = ""
                Dim TeachID_1 As String = ""
                Dim Selstr2 As String = ""
                Dim ClassID_2 As String = ""
                Dim TeachID_2 As String = ""
                'Dim i As Integer=0
                Dim OldDt As New DataTable  '變更前
                Dim NewDt As New DataTable  '變更後 
                Dim ClassRow1 As DataRow
                Dim ClassRow2 As DataRow
                Dim Sqlnew As String = "" '原始排課資料
                Dim Sqlold As String = "" '變更排課資料
                Dim Sqlnew2 As String = "" '原始排課資料
                Dim Sqlold2 As String = "" '變更排課資料

                If TargetLB1.Items.Count <> TargetLB2.Items.Count Then
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "變更課程項目堂數需一致!!")
                    Exit Sub
                End If
                If TargetLB2.Items.Count = 0 Then              '判斷變更計畫是否有選課程
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "變更計畫內容請選擇課程!!!")
                    Exit Sub
                End If
                If TargetLB1.Items.Count = 0 Then              '判斷原計畫是否有選課程
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "原計畫內容請選擇課程!-CACD")
                    Exit Sub
                End If

                'Dim objconn As SqlConnection=DbAccess.GetConnection()
                'Call TIMS.OpenDbConn(conn)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    '(檢查資料)
                    OldDt.Columns.Add(New DataColumn("orderid"))  '節次
                    OldDt.Columns.Add(New DataColumn("Class"))
                    OldDt.Columns.Add(New DataColumn("Room"))
                    OldDt.Columns.Add(New DataColumn("Teacher1"))
                    OldDt.Columns.Add(New DataColumn("Teacher2"))
                    NewDt.Columns.Add(New DataColumn("orderid"))  '節次
                    NewDt.Columns.Add(New DataColumn("Class"))
                    NewDt.Columns.Add(New DataColumn("Room"))
                    NewDt.Columns.Add(New DataColumn("Teacher1"))
                    NewDt.Columns.Add(New DataColumn("Teacher2"))
                    ClassRow1 = Get_ClassRow(ViewState(vs_OCID), CDate(TimeSDate.Text).ToString("yyyy/MM/dd"), TransConn, Trans) '取得原計畫排課資料
                    ClassRow2 = Get_ClassRow(ViewState(vs_OCID), CDate(TimeEDate.Text).ToString("yyyy/MM/dd"), TransConn, Trans) '取得變更計畫排課資料 'Get_ClassRow(TimeEDate, Sqlda)
                    If hid_chklist.Value = "N" Then
                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        Common.MessageBox(Page, "變更申請失敗!欲更換日期兩日皆無安排課程！!")
                        Exit Sub
                    End If

                    For i As Integer = 0 To TargetLB1.Items.Count - 1
                        Dim dr4 As DataRow
                        dr4 = OldDt.NewRow
                        If TargetLB1.Items(i).Value <> "" Then
                            dr4("orderid") = TargetLB1.Items(i).Value                                '節次
                            dr4("Class") = ClassRow1("class" & TargetLB1.Items(i).Value)             '課程
                            dr4("Room") = ClassRow1("Room" & TargetLB1.Items(i).Value)               '教室
                            dr4("Teacher1") = ClassRow1("Teacher" & TargetLB1.Items(i).Value)        '教師(一)
                            dr4("Teacher2") = ClassRow1("Teacher" & Convert.ToString(CInt(TargetLB1.Items(i).Value) + 12).ToString)    '教師(二)
                            OldDt.Rows.Add(dr4)
                            '取得原計畫節次
                            If Selstr1 <> "" Then Selstr1 &= ","
                            Selstr1 &= TargetLB1.Items(i).Value
                            '取得原計畫課程 classid 
                            If ClassID_1 <> "" Then ClassID_1 &= ","
                            ClassID_1 &= If(Trim(Convert.ToString(ClassRow1("class" & TargetLB1.Items(i).Value))) = "", "x", Trim(Convert.ToString(ClassRow1("class" & TargetLB1.Items(i).Value))))
                            '取得原計畫 teacherid 
                            If TeachID_1 <> "" Then TeachID_1 &= ","
                            TeachID_1 &= If(Trim(Convert.ToString(ClassRow1("Teacher" & TargetLB1.Items(i).Value))) = "", "x", Trim(Convert.ToString(ClassRow1("Teacher" & TargetLB1.Items(i).Value))))
                        End If
                    Next

                    For i As Integer = 0 To TargetLB2.Items.Count - 1
                        Dim dr5 As DataRow
                        dr5 = NewDt.NewRow
                        If Trim(TargetLB2.Items(i).Value) <> "" Then
                            dr5("orderid") = TargetLB2.Items(i).Value                                 '節次
                            dr5("Class") = ClassRow2("class" & TargetLB2.Items(i).Value)              '課程
                            dr5("Room") = ClassRow2("Room" & TargetLB2.Items(i).Value)                '教室
                            dr5("Teacher1") = ClassRow2("Teacher" & TargetLB2.Items(i).Value)         '教師(一)
                            dr5("Teacher2") = ClassRow2("Teacher" & Convert.ToString(CInt(TargetLB2.Items(i).Value) + 12).ToString)    '教師(二)
                            NewDt.Rows.Add(dr5)
                            '取得變更計畫節次
                            If Selstr2 <> "" Then Selstr2 &= ","
                            Selstr2 &= TargetLB2.Items(i).Value
                            '取得變更計畫課程 classid 
                            If ClassID_2 <> "" Then ClassID_2 &= ","
                            ClassID_2 &= If(Trim(Convert.ToString(ClassRow2("class" & TargetLB2.Items(i).Value))) = "", "x", Trim(Convert.ToString(ClassRow2("class" & TargetLB2.Items(i).Value))))
                            '取得變更計畫 teacherid 
                            If TeachID_2 <> "" Then TeachID_2 &= ","
                            TeachID_2 &= If(Trim(Convert.ToString(ClassRow2("Teacher" & TargetLB2.Items(i).Value))) = "", "x", Trim(Convert.ToString(ClassRow2("Teacher" & TargetLB2.Items(i).Value))))
                        End If
                    Next

                    '檢查array是否一致
                    Dim aryClass1 As Array = Split(ClassID_1, ",")
                    Dim aryClass2 As Array = Split(ClassID_2, ",")
                    Dim aryTeachID1 As Array = Split(TeachID_1, ",")
                    Dim aryTeachID2 As Array = Split(TeachID_2, ",")

                    If aryClass1.Length <> aryClass1.Length Then
                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        Common.MessageBox(Page, "變更申請失敗!課程數不相符!")
                        Exit Sub
                    End If
                    If aryTeachID1.Length <> aryTeachID2.Length Then
                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        Common.MessageBox(Page, "變更申請失敗!師資數不相符!")
                        Exit Sub
                    End If
                    If aryClass1.Length <> aryTeachID1.Length Then
                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        Common.MessageBox(Page, "變更申請失敗!師資及課程數不相符!")
                        Exit Sub
                    End If

                    '檢查是否與審核中的課程與要申請的課程是否有重覆 
                    Dim SDateExistClass As String = ChkIsClassMatch(Selstr1, TimeSDate.Text, TransConn, Trans)
                    Dim EDateExistClass As String = ChkIsClassMatch(Selstr2, TimeEDate.Text, TransConn, Trans)
                    If SDateExistClass <> "" Then
                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        'Common.MessageBox(Page, "變更申請失敗！" & Convert.ToDateTime(TimeSDate.Text).ToString("yyyy-MM-dd") & " 欲變更之日期節次與" & SDateExistClass & " 申請之變更有重疊，請協調中心承辦人先進行未審核之變更申請的審核動作!!")
                        Common.MessageBox(Page, "變更申請失敗！" & Convert.ToDateTime(TimeSDate.Text).ToString("yyyy-MM-dd") & " 欲變更之日期節次與" & SDateExistClass & " 申請之變更有重疊，請協調分署承辦人先進行未審核之變更申請的審核動作!!")
                        Exit Sub
                    End If
                    If EDateExistClass <> "" Then
                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        'Common.MessageBox(Page, "變更申請失敗！" & Convert.ToDateTime(TimeEDate.Text).ToString("yyyy-MM-dd") & " 欲變更之日期節次與" & EDateExistClass & " 申請之變更有重疊，請協調中心承辦人先進行未審核之變更申請的審核動作!!")
                        Common.MessageBox(Page, "變更申請失敗！" & Convert.ToDateTime(TimeEDate.Text).ToString("yyyy-MM-dd") & " 欲變更之日期節次與" & EDateExistClass & " 申請之變更有重疊，請協調分署承辦人先進行未審核之變更申請的審核動作!!")
                        Exit Sub
                    End If

                    '(存入異動申請項目) 'Exit Function
                    sql = ""
                    sql &= " INSERT INTO PLAN_REVISE (PlanID ,ComIDNO ,SeqNO ,SubSeqNo ,CDate ,AltDataID" & vbCrLf
                    sql &= "  ,OldData2_1 ,OldData2_2 ,OldData2_3 ,NewData2_1 ,NewData2_2 ,NewData2_3" & vbCrLf
                    sql &= "  ,ReviseAcct ,ReviseCont ,changeReason" & vbCrLf
                    sql &= "  ,Verifier ,ReviseStatus ,ModifyAcct ,ModifyDate)" & vbCrLf
                    sql &= " VALUES (@PlanID ,@ComIDNO, @SeqNO ,@SubSeqNo ,convert(date,@CDate),@AltDataID" & vbCrLf
                    sql &= "  ,convert(date,@OldData2_1) ,@OldData2_2 ,@OldData2_3" & vbCrLf
                    sql &= "  ,convert(date,@NewData2_1) ,@NewData2_2 ,@NewData2_3" & vbCrLf
                    sql &= "  ,@ReviseAcct ,@ReviseCont ,@changeReason" & vbCrLf
                    sql &= "  ,@Verifier ,@ReviseStatus ,@ModifyAcct ,GETDATE())" & vbCrLf

                    Dim i_Parms As New Hashtable
                    i_Parms.Add("PlanID", rPlanID) 'Request("PlanID")
                    i_Parms.Add("ComIDNO", rComIDNO) 'Request("cid")
                    i_Parms.Add("SeqNO", rSeqNo) 'Request("no")
                    i_Parms.Add("SubSeqNo", iSubSeqNO)
                    i_Parms.Add("CDate", TIMS.Cdate3(ApplyDate.Text))
                    i_Parms.Add("AltDataID", v_ChgItem) 'ChgItem.SelectedValue
                    '異動項目
                    i_Parms.Add("OldData2_1", TIMS.Cdate3(TimeSDate.Text)) ' CDate(TimeSDate.Text)
                    i_Parms.Add("OldData2_2", ClassID_1) '課程
                    i_Parms.Add("OldData2_3", Selstr1) '節次，逗號分開
                    i_Parms.Add("NewData2_1", TIMS.Cdate3(TimeEDate.Text)) 'CDate(TimeEDate.Text)
                    i_Parms.Add("NewData2_2", ClassID_2) '課程
                    i_Parms.Add("NewData2_3", Selstr2) '節次，逗號分開

                    i_Parms.Add("ReviseAcct", sm.UserInfo.UserID)
                    i_Parms.Add("ReviseCont", If(ReviseCont.Text <> "", ReviseCont.Text, Convert.DBNull))
                    i_Parms.Add("changeReason", If(v_changeReason <> "", v_changeReason, Convert.DBNull))
                    i_Parms.Add("Verifier", If(iPlanKind = 1, sm.UserInfo.UserID, Convert.DBNull))
                    i_Parms.Add("ReviseStatus", If(iPlanKind = 1, "Y", Convert.DBNull))
                    i_Parms.Add("ModifyAcct", sm.UserInfo.UserID)

                    Try
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(sql, Trans, i_Parms)

                    Catch ex As Exception
                        Dim strErrmsg5 As String = ""
                        strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                        strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                        strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                        strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                        strErrmsg5 &= "/* sql: */" & vbCrLf
                        strErrmsg5 &= sql & vbCrLf
                        Call TIMS.WriteTraceLog(strErrmsg5)

                        DbAccess.RollbackTrans(Trans)
                        Call TIMS.CloseDbConn(TransConn)
                        If flagDebugTest Then Throw ex
                    End Try

                    'Sqlda.SelectCommand.CommandText=sql
                    'Sqlda.SelectCommand.ExecuteNonQuery()
                    '(存入異動申請項目)
                    '20080912 andy edit
                    If iPlanKind = 1 Then  '自辦時
                        Sqlnew = ""
                        Sqlold = ""
                        Sqlnew = " UPDATE CLASS_SCHEDULE SET ModifyAcct='" & sm.UserInfo.UserID & "' ,ModifyDate=GETDATE()," & vbCrLf
                        Sqlold = " UPDATE CLASS_SCHEDULE SET ModifyAcct='" & sm.UserInfo.UserID & "' ,ModifyDate=GETDATE()," & vbCrLf

                        Sqlnew2 = ""
                        Sqlold2 = ""
                        For m1 As Int16 = 0 To NewDt.Rows.Count - 1 'NEW ROW
                            Dim dr01 As DataRow = OldDt.Rows(m1) ' OLD ROW
                            Dim dr02 As DataRow = NewDt.Rows(m1) ' NEW ROW
                            '更新 Teacher id 講師 
                            Dim sTmp1 As String = ""
                            '講師(一)
                            sTmp1 = "Teacher" & Convert.ToString(dr01("orderid")) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr02("Teacher1")))
                            Sqlnew2 = TIMS.GetCommaStr(Sqlnew2, sTmp1)
                            '講師(二)
                            sTmp1 = "Teacher" & Convert.ToString(CInt(dr01("orderid")) + 12) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr02("Teacher2")))
                            Sqlnew2 = TIMS.GetCommaStr(Sqlnew2, sTmp1)
                            '更新 class id 課程 
                            sTmp1 = "Class" & Convert.ToString(dr01("orderid")) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr02("Class")))
                            Sqlnew2 = TIMS.GetCommaStr(Sqlnew2, sTmp1)
                            '更新 Room id 教室(上課地點) 
                            sTmp1 = "Room" & Convert.ToString(dr01("orderid")) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr02("Room")))
                            Sqlnew2 = TIMS.GetCommaStr(Sqlnew2, sTmp1)
                            '更新 Teacher id 講師
                            '講師(一)
                            sTmp1 = "Teacher" & Convert.ToString(dr02("orderid")) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr01("Teacher1")))
                            Sqlold2 = TIMS.GetCommaStr(Sqlold2, sTmp1)
                            '講師(二)
                            sTmp1 = "Teacher" & Convert.ToString(CInt(dr02("orderid")) + 12) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr01("Teacher2")))
                            Sqlold2 = TIMS.GetCommaStr(Sqlold2, sTmp1)
                            '更新 class id 課程 
                            sTmp1 = "Class" & Convert.ToString(dr02("orderid")) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr01("Class")))
                            Sqlold2 = TIMS.GetCommaStr(Sqlold2, sTmp1)
                            '更新 Room id 教室(上課地點)
                            sTmp1 = "Room" & Convert.ToString(dr02("orderid")) & "=" & TIMS.Get_NULLvalue(Convert.ToString(dr01("Room")))
                            Sqlold2 = TIMS.GetCommaStr(Sqlold2, sTmp1)
                            'For m2 As Int16=0 To OldDt.Rows.Count - 1
                            '    dr01=OldDt.Rows(m2)
                            '    If m1=m2 Then
                            '    End If
                            'Next
                        Next
                        Sqlnew = Trim(Sqlnew)
                        Sqlnew = Sqlnew + Sqlnew2
                        Sqlnew = Sqlnew & " WHERE OCID=" & ViewState(vs_OCID)
                        'Sqlnew=Sqlnew & " AND SchoolDate='" & TimeSDate.Text & "'" & vbCrLf
                        Sqlnew = Sqlnew & "  AND SchoolDate=" & TIMS.To_date(TimeSDate.Text) & "" & vbCrLf

                        Sqlold = Trim(Sqlold)
                        Sqlold = Sqlold + Sqlold2
                        Sqlold = Sqlold & " WHERE OCID=" & ViewState(vs_OCID)
                        'Sqlold=Sqlold & " AND SchoolDate='" & TimeEDate.Text & "'" & vbCrLf
                        Sqlold = Sqlold & "  AND SchoolDate=" & TIMS.To_date(TimeEDate.Text) & "" & vbCrLf

                        If DateDiff(DateInterval.Day, CDate(TimeSDate.Text), CDate(TimeEDate.Text)) = 0 Then
                            DbAccess.ExecuteNonQuery(Sqlnew, Trans)
                        Else
                            'DbAccess.ExecuteNonQuery(Sqlnew & Sqlold, objTrans)
                            DbAccess.ExecuteNonQuery(Sqlnew, Trans)
                            DbAccess.ExecuteNonQuery(Sqlold, Trans)
                        End If
                        'If TimeSDate.Text=TimeEDate.Text Then
                        '    Sqlda.SelectCommand.CommandText=Sqlnew
                        'Else
                        '    Sqlda.SelectCommand.CommandText=Sqlnew & Sqlold
                        'End If
                        'Sqlda.SelectCommand.ExecuteNonQuery(sql, Trans)
                    End If
                    DbAccess.CommitTrans(Trans)
                    'TIMS.CloseDbConn(objconn)
                    'If conn.State=ConnectionState.Open Then conn.Close()
                    'Trans.Commit()
                    'Sqlda.Dispose()
                    If Not OldDt Is Nothing Then OldDt.Dispose()
                    If Not NewDt Is Nothing Then NewDt.Dispose()
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    strErrmsg5 &= "/* Sqlnew: */" & vbCrLf
                    strErrmsg5 &= Sqlnew & vbCrLf
                    strErrmsg5 &= "/* Sqlold: */" & vbCrLf
                    strErrmsg5 &= Sqlold & vbCrLf
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-B119")
                    Dim vErrMsg1 As String = ""
                    vErrMsg1 &= "【發生錯誤】" & vbCrLf
                    vErrMsg1 &= ex.ToString & vbCrLf
                    Common.MessageBox(Page, vErrMsg1)
                    'Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try
                strMassage = TIMS.IsAllDateCheck(Me, ViewState(vs_OCID), "ReturnMsg", gobjconn) '檢查是否為全日制,若是全日制檢查是否符合規則

            Case Cst_i訓練地點 'As Integer=3
                '變更教室
                '(檢查資料)
                '(TIMS)訓練地點
                'Dim i As Integer
                Dim str As String = ""
                For i As Integer = 0 To SPlace.Items.Count - 1
                    If SPlace.Items(i).Selected Then str += SPlace.Items(i).Value.Replace(" Then,", ";") & ","
                Next
                'Dim Onetmp()
                If str = "" Then str = "a"
                Dim Onetmp() As String = Split(Left(str, Len(str) - 1), ",")
                Dim Onetmp2(2) As String
                TotalItem = ""
                For i As Integer = 0 To Onetmp.Length - 1
                    Onetmp2 = Split(Onetmp(i), "^")
                    TotalItem += Onetmp2(0) & ","
                    If Onetmp2.Length > 1 Then
                        If i = 0 Then
                            campare = Onetmp2(1)
                        Else
                            If campare <> Onetmp2(1) Then
                                Common.MessageBox(Page, "原計畫內容請選擇相同地點!!!")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                If TotalItem = "," Then           '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "原計畫內容請選擇課程!-BEAE")
                    Exit Sub
                End If
                If EPlace.Text = "" Then         '判斷是否有填更換地點
                    Common.MessageBox(Page, "請填入要更換地點!!!")
                    Exit Sub
                End If
                If ReviseCont.Text.Length > 255 Then
                    Common.MessageBox(Page, "變更原因字串長度過長255!!!")
                    Exit Sub
                End If

                '(存入異動申請項目)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練地點, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData3_1") = PlaceDate.Text
                    dr("OldData3_2") = campare         '(TIMS)訓練地點
                    dr("OldData3_3") = Left(TotalItem, Len(TotalItem) - 1)      '節次，逗號分開
                    dr("NewData3_1") = EPlace.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-4EA1")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    'sql=" Select * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND SchoolDate='" & PlaceDate.Text & "'"
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND SchoolDate=" & TIMS.To_date(PlaceDate.Text)
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim ClassNum As Array = Split(Left(TotalItem, Len(TotalItem) - 1), ",")
                        For j As Integer = 0 To ClassNum.Length - 1
                            dr("Room" & ClassNum(j)) = EPlace.Text
                        Next
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i課程編配 'As Integer=4
                '更變訓練時數
                '(檢查資料)
                If EGenSci.Text = "" Then           '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "請輸入一般學科!!!")
                    Exit Sub
                End If
                If EProSci.Text = "" Then           '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "請輸入一般術科!!!")
                    Exit Sub
                End If
                If EProTech.Text = "" Then           '判斷原計畫是否有選課程
                    Common.MessageBox(Page, "請輸入術科!!!")
                    Exit Sub
                End If

                '(存入異動申請項目)
                'Dim conn As SqlConnection=DbAccess.GetConnection()
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i課程編配, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData4_1") = SSumSci.Text
                    dr("OldData4_2") = SGenSci.Text
                    dr("OldData4_3") = SProSci.Text
                    dr("OldData4_4") = SProTech.Text
                    dr("OldData4_5") = SOther.Text
                    dr("NewData4_1") = ESumSci.Text
                    dr("NewData4_2") = EGenSci.Text
                    dr("NewData4_3") = EProSci.Text
                    dr("NewData4_4") = EProTech.Text
                    dr("NewData4_5") = EOther.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)

                    Common.MessageBox(Page, "計畫變更儲存失敗!-5737")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    sql = " Select * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("GenSciHours") = EGenSci.Text
                        dr("ProSciHours") = EProSci.Text
                        dr("ProTechHours") = EProTech.Text
                        dr("OtherHours") = EOther.Text
                        dr("TotalHours") = TIMS.CINT1(EGenSci.Text) + TIMS.CINT1(EProSci.Text) + TIMS.CINT1(EProTech.Text) + TIMS.CINT1(EOther.Text) '20060525 by Vicient
                        dr("Thours") = TIMS.CINT1(EGenSci.Text) + TIMS.CINT1(EProSci.Text) + TIMS.CINT1(EProTech.Text) + TIMS.CINT1(EOther.Text)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                        '20060525 by Vicient start
                        sql = " SELECT * FROM CLASS_CLASSINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                        dt = DbAccess.GetDataTable(sql, da, TransConn)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("THours") = TIMS.CINT1(EGenSci.Text) + TIMS.CINT1(EProSci.Text) + TIMS.CINT1(EProTech.Text) + TIMS.CINT1(EOther.Text)
                            dr("LastState") = "M" 'M: 修改(最後異動狀態)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da)
                        End If
                        'end
                    End If
                End If

            Case Cst_i訓練師資 'As Integer=5 訓練師資
                '(檢查資料)
                'Dim i As Integer
                Dim str As String = ""
                '節次
                For i As Integer = 0 To STeacher.Items.Count - 1
                    '要變更資料組合
                    If STeacher.Items(i).Selected AndAlso STeacher.Items(i).Value <> "" Then str += STeacher.Items(i).Value & ";"   'ex:  "2^54273,54278;" --> "TecherID1,(第二筆)TecherID2"  ※ TecherID2=TecherID1+12(欄位)
                Next
                'str: 1^6480,45754;2^6480,45754; 1.2為變更前節次，後為變更前師資
                If str = "" Then
                    Common.MessageBox(Page, "請於原課程資料中，選擇師資資料進行變更。")
                    Exit Sub
                End If
                Dim Onetmp As String()
                If str = "" Then str = "a"
                Onetmp = Split(Left(str, Len(str) - 1), ";") '要變更資料組合
                'Dim Onetmp2(2)
                TotalItem = ""
                'TotalItem 
                hid_NoTechID2.Value = "" '含有無助教1的資料 (師資2)
                hid_NoTechID3.Value = "" '含有無助教2的資料
                For i As Integer = 0 To Onetmp.Length - 1
                    'str: 1^6480,45754;2^6480,45754; 1.2為變更前節次，後為變更前師資
                    Dim Onetmp2 As String() = Split(Onetmp(i), "^") '課堂/師資1/助教1/助教2 
                    TotalItem += Onetmp2(0) & "," '節次
                    If Onetmp2.Length > 1 Then
                        If i = 0 Then
                            campare = Onetmp2(1) '變更內容 第1組加入(TECH,ROOM,CLASS)
                        Else
                            '後續每組比較
                            If campare <> Onetmp2(1) Then
                                Common.MessageBox(Page, "請於原課程資料中，選擇相同師資進行變更。")
                                Exit Sub
                            End If
                        End If
                        Dim Tchtmp3 As String() = Split(Onetmp2(1), ",")
                        '含有無師資2的資料
                        '有2組
                        hid_NoTechID2.Value = "Y"
                        If Tchtmp3.Length > 1 Then
                            If "" & Convert.ToString(Tchtmp3(1)) <> "" Then hid_NoTechID2.Value = "" '有資料，清除判斷
                        End If
                        '有3組
                        hid_NoTechID3.Value = "Y"
                        If Tchtmp3.Length > 2 Then
                            If "" & Convert.ToString(Tchtmp3(2)) <> "" Then hid_NoTechID3.Value = "" '有資料，清除判斷
                        End If
                    End If
                    'If Onetmp2.Length > 2 Then
                    '    If i=0 Then
                    '        campare2=Onetmp2(2)
                    '    Else
                    '        If campare2 <> Onetmp2(2) Then
                    '            Common.MessageBox(Page, "請於原課程資料中，選擇相同師資2進行變更。")
                    '            Exit Function
                    '        End If
                    '    End If
                    'End If
                Next
                '非自辦計畫使用
                If TIMS.Cst_TPlanID02Plan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                    If OLessonTeah2Value.Value <> "" AndAlso hid_NoTechID2.Value = "Y" Then
                        Common.MessageBox(Page, Cst_msg2)
                        Exit Sub
                    End If
                    If OLessonTeah3Value.Value <> "" AndAlso hid_NoTechID3.Value = "Y" Then
                        Common.MessageBox(Page, Cst_msg3)
                        Exit Sub
                    End If
                End If

                '(存入異動申請項目)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練師資, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData5_1") = TechDate.Text
                    dr("OldData5_2") = campare          '師資
                    dr("OldData5_3") = Left(TotalItem, Len(TotalItem) - 1)      '節次，逗號分開
                    '20081125 andy edit 修改師資為師資(一)、師資(二) *若有資料(二)在審核時值寫入  師資(一)+12的欄位 內
                    'dr("NewData5_1")=OLessonTeah1Value.Value
                    Dim ssNewData5_1 As String = ""
                    If ssNewData5_1 = "" AndAlso OLessonTeah3Value.Value <> "" Then
                        ssNewData5_1 = OLessonTeah1Value.Value
                        ssNewData5_1 &= "," & OLessonTeah2Value.Value
                        ssNewData5_1 &= "," & OLessonTeah3Value.Value
                    End If
                    If ssNewData5_1 = "" AndAlso OLessonTeah2Value.Value <> "" Then
                        ssNewData5_1 = OLessonTeah1Value.Value
                        ssNewData5_1 &= "," & OLessonTeah2Value.Value
                    Else
                        ssNewData5_1 = OLessonTeah1Value.Value
                    End If
                    dr("NewData5_1") = ssNewData5_1

                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-7755")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    '自辦直接變更原始資料
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND SchoolDate=" & TIMS.To_date(TechDate.Text)
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim ClassNum As Array = Split(Left(TotalItem, Len(TotalItem) - 1), ",")
                        Dim ssSave1 As Boolean = False
                        If Not ssSave1 AndAlso OLessonTeah3Value.Value <> "" Then
                            ssSave1 = True
                            For j As Integer = 0 To ClassNum.Length - 1
                                dr("Teacher" & ClassNum(j)) = OLessonTeah1Value.Value
                                dr("Teacher" & ClassNum(j) + 12) = OLessonTeah2Value.Value
                                dr("Teacher" & ClassNum(j) + 24) = OLessonTeah3Value.Value
                            Next
                        End If
                        If Not ssSave1 AndAlso OLessonTeah2Value.Value <> "" Then
                            For j As Integer = 0 To ClassNum.Length - 1
                                dr("Teacher" & ClassNum(j)) = OLessonTeah1Value.Value
                                dr("Teacher" & ClassNum(j) + 12) = OLessonTeah2Value.Value
                            Next
                        End If
                        If Not ssSave1 AndAlso OLessonTeah1Value.Value <> "" Then
                            For j As Integer = 0 To ClassNum.Length - 1
                                dr("Teacher" & ClassNum(j)) = OLessonTeah1Value.Value
                            Next
                        End If
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i班別名稱 'As Integer=6
                '(檢查資料)
                If ChangeClassCName.Text = "" Then         '判斷是否有填班別名稱
                    Common.MessageBox(Page, "請填入要更換班別名稱!!!")
                    Exit Sub
                End If

                '(存入異動申請項目)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i班別名稱, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData6_1") = ClassCName.Text
                    dr("NewData6_1") = ChangeClassCName.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-E5A1")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    sql = " Select * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("ClassName") = ChangeClassCName.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("ClassCName") = ChangeClassCName.Text
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i期別 'As Integer=7
                '(檢查資料)
                Dim sPMS As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNO", rSeqNo}}
                Dim objstr As String = ""
                objstr &= " Select a.CyclType ,c.RID "
                objstr &= " FROM PLAN_PLANINFO a "
                objstr &= " JOIN ORG_ORGINFO b On a.ComIDNO=b.ComIDNO "
                objstr &= " JOIN AUTH_RELSHIP c On b.OrgID=c.OrgID "
                objstr &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO"
                Dim objrow As DataRow = DbAccess.GetOneRow(objstr, gobjconn, sPMS)

                Dim csPMS As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNO", rSeqNo}}
                chkstr = "" 'Dim chkstr As String="" 'Dim objrow2 As DataRow=Nothing
                chkstr &= " SELECT * FROM PLAN_PLANINFO "
                chkstr &= " WHERE TransFlag='Y' AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO =@SeqNO"
                Dim objrow2 As DataRow = DbAccess.GetOneRow(chkstr, gobjconn, csPMS)

                Dim sPMS3 As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNO", rSeqNo}}
                Dim objstr3 As String = ""
                objstr3 &= " SELECT b.CLSID ,b.YEARS ,c.CLASSID ,b.CYCLTYPE,b.OCID,b.CLASSCNAME" & vbCrLf
                objstr3 &= " FROM PLAN_PLANINFO a" & vbCrLf
                objstr3 &= " JOIN CLASS_CLASSINFO b ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNO=b.SeqNO" & vbCrLf
                objstr3 &= " JOIN ID_Class c ON b.CLSID=c.CLSID" & vbCrLf
                objstr3 &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO"
                Dim objrow3 As DataRow = DbAccess.GetOneRow(objstr3, gobjconn, sPMS3)

                ChangeCyclType.Text = TIMS.FmtCyclType(ChangeCyclType.Text)

                Dim chkstr3 As String = ""
                Dim chkstr4 As String = ""
                If objrow3 IsNot Nothing Then
                    '變更的期別與開班資料中的期別重覆-在職
                    chkstr3 = ""
                    chkstr3 &= " SELECT 'x' "
                    chkstr3 &= " FROM CLASS_CLASSINFO cc"
                    chkstr3 &= " WHERE cc.CLSID='" & objrow3("CLSID") & "'"
                    chkstr3 &= " AND cc.PlanID=" & rPlanID & vbCrLf
                    chkstr3 &= " AND cc.CyclType='" & ChangeCyclType.Text & "'"
                    chkstr3 &= " AND cc.RID='" & objrow("RID") & "'"
                    chkstr3 &= " AND cc.CLASSCNAME='" & objrow3("CLASSCNAME") & "'"
                    chkstr3 &= " AND cc.OCID != '" & objrow3("OCID") & "'"

                    chkstr4 = ""
                    chkstr4 &= " SELECT 'x' "
                    chkstr4 &= " FROM CLASS_STUDENTSOFCLASS cs "
                    chkstr4 &= " JOIN CLASS_CLASSINFO cc ON cc.ocid=cs.ocid "
                    chkstr4 &= " WHERE cc.PLANID=" & rPlanID 'Request("PlanID")  
                    chkstr4 &= " AND cc.OCID ='" & objrow3("OCID") & "'" 'Request("PlanID")  
                End If
                If ChangeCyclType.Text = "" Then         '判斷是否有填期別
                    Common.MessageBox(Page, "請填入要更換期別!!!")
                    Exit Sub
                End If
                If ChangeCyclType.Text.Length <> 2 Then
                    Common.MessageBox(Page, "期別要輸入2位數字才行!!!")
                    Exit Sub
                End If
                If ChangeCyclType.Text = CyclType.Text Then
                    Common.MessageBox(Page, "變更的期別不可和原期別相同!!!")
                    Exit Sub
                End If
                If chkstr3 <> "" AndAlso DbAccess.GetCount(chkstr3, gobjconn) > 0 Then
                    Common.MessageBox(Page, "變更的期別與開班資料中的期別重覆!!!")
                    Exit Sub
                End If
                'Common.MessageBox(Page, "沒有轉班才可改!!!")
                If objrow2 IsNot Nothing AndAlso chkstr4 <> "" AndAlso DbAccess.GetCount(chkstr4, gobjconn) > 0 Then
                    Common.MessageBox(Page, "此計劃班別已有學員資料，不可變更期別!!!")
                    Exit Sub
                End If

                '(存入異動申請項目)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i期別, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    dr("OldData6_1") = ClassCName2.Text
                    dr("OldData7_1") = TIMS.FmtCyclType(CyclType.Text)
                    dr("NewData7_1") = ChangeCyclType.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-0843")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim vCyclType As String = TIMS.FmtCyclType(ChangeCyclType.Text)
                        dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        Dim vCyclType As String = TIMS.FmtCyclType(ChangeCyclType.Text)
                        dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i上課地址 'As Integer=8 變更訓練地點
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課地址, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    dr("OldData8_1") = If(OldData8_1.Value = "", Convert.DBNull, OldData8_1.Value)
                    dr("OldData8_3") = If(OldData8_3.Value = "", Convert.DBNull, OldData8_3.Value)
                    dr("OldData8_2") = If(OldData8_2.Value = "", Convert.DBNull, OldData8_2.Value)

                    dr("NewData8_1") = NewData8_1.Value
                    dr("NewData8_3") = If(NewData8_3.Value <> "", NewData8_3.Value, Convert.DBNull)
                    dr("NewData8_2") = NewData8_2.Text

                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-B765")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
                If iPlanKind = 1 Then
                    sql = " Select * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TaddressZip") = NewData8_1.Value
                        dr("TaddressZIP6W") = If(hidNewData8_6W.Value <> "", hidNewData8_6W.Value, Convert.DBNull)
                        dr("TAddress") = NewData8_2.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TaddressZip") = NewData8_1.Value
                        dr("TaddressZIP6W") = If(hidNewData8_6W.Value <> "", hidNewData8_6W.Value, Convert.DBNull)
                        dr("TAddress") = NewData8_2.Text
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i停辦 'As Integer=9 '變更停辦狀態
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i停辦, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData9_1") = "N"
                    dr("NewData9_1") = "Y"
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-609D")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try
            ''不開班原因存入 amu 20061
            'Try
            '    Dim sql2 As String
            '    Dim dt2 As DataTable
            '    Dim dr2 As DataRow
            '    sql2="Select * FROM CLASS_CLASSINFO WHERE PlanID='" & rPlanID 'Request("PlanID") & "' and ComIDNO='" & rComIDNO 'Request("cid") & "' and SeqNo='" & rSeqNo 'Request("no") & "'"
            '    dt2=DbAccess.GetDataTable(sql2)
            '    If dt2.Rows.Count > 0 Then
            '        dr2=dt2.Rows(0)
            '        dr2("NotOpen")="Y"
            '        dr2("NORID")=""
            '        'For i As Integer=0 To NORID.Items.Count - 1
            '        '    If NORID.Items(i).Selected=True Then
            '        '    If dr2("NORID")="" Then
            '        '        dr2("NORID")=NORID.Items(i).Value
            '        '    Else
            '        '        dr2("NORID") += "," & NORID.Items(i).Value
            '        '    End If
            '        'End If
            '        'Next
            '        dr2("OtherReason")=If(ReviseCont.Text="", Convert.DBNull, ReviseCont.Text)
            '        dr2("LastState")="D" 'D: 刪除(最後異動狀態)
            '    End If
            'Catch ex As Exception
            '    If flagDebugTest Then Throw ex
            'End Try

            Case Cst_i上課時段 'As Integer=10
                '(存入申請異動項目)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課時段, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData10_1") = OldData10_1.Value
                    dr("NewData10_1") = NewData10_1.SelectedValue
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-B451")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i師資 'As Integer=11 '變更師資
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i師資, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData11_1") = OldData11_1.Value
                    dr("OldData11_2") = TeacherName1.Text
                    dr("NewData11_1") = NewData11_1.Value
                    dr("NewData11_2") = TeacherName1_2.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-ADD6")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i助教 'As Integer=20  '20120213 BY AMU (產投用助教) '變更助教
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    'Trans=DbAccess.BeginTrans(Transconn)
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i助教, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    dr("OldData20_1") = OldData20_1.Value
                    dr("OldData20_2") = TeacherName2.Text
                    dr("NewData20_1") = NewData20_1.Value
                    dr("NewData20_2") = TeacherName2_2.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-B53C")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i訓練費用 'As Integer=21  '20170908 (職前)  '變更訓練費用
                '(存入申請異動項目)
                If Session(Hid_COSTITEM_GUID21.Value) Is Nothing Then
                    Common.MessageBox(Page, "計畫變更儲存，查無資料!!")
                    Exit Sub
                End If
                Dim dtC21 As DataTable = Session(Hid_COSTITEM_GUID21.Value)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i訓練費用, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    Dim dtTemp2 As DataTable = Session(Hid_COSTITEM_GUID21.Value)

                    Dim htSEL As New Hashtable
                    TIMS.SetMyValue2(htSEL, "rPlanID", Convert.ToString(rPlanID)) '計畫PK
                    TIMS.SetMyValue2(htSEL, "rComIDNO", rComIDNO) '計畫PK
                    TIMS.SetMyValue2(htSEL, "rSeqNo", Convert.ToString(rSeqNo)) '計畫PK
                    TIMS.SetMyValue2(htSEL, "rApplyDate", TIMS.Cdate3(ApplyDate.Text)) 'ApplyDate.Text
                    TIMS.SetMyValue2(htSEL, "SubSeqNO", iSubSeqNO) 'iSubSeqNO
                    TIMS.SetMyValue2(htSEL, "CostMode", Hid_CostMode.Value) 'CostMode
                    TIMS.SetMyValue2(htSEL, "ssUserID", Convert.ToString(sm.UserInfo.UserID))
                    Call SAVE_REVISE_COSTITEM(htSEL, dtTemp2, TransConn, Trans)

                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-864C")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i核定人數 'As Integer=12  'Cst_招生人數  'As Integer=12 '變更招生人數
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    'Trans=DbAccess.BeginTrans(Transconn)
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i核定人數, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    dr("OldData12_1") = OldData12_1.Text
                    dr("NewData12_1") = NewData12_1.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    If iPlanKind = 1 Then
                        sql = " Select * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("TNum") = NewData12_1.Text
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da, Trans)
                        End If
                        sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("TNum") = NewData12_1.Text
                            dr("LastState") = "M" 'M: 修改(最後異動狀態)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da, Trans)
                        End If
                    End If
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-9CA7")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TNum") = NewData12_1.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                    sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    '2006/03/ add conn by matt
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("TNum") = NewData12_1.Text
                        dr("LastState") = "M" 'M: 修改(最後異動狀態)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i增班 'As Integer=13 '變更班數
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    'Trans=DbAccess.BeginTrans(Transconn)
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i增班, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    dr("OldData13_1") = OldData13_1.Text
                    dr("NewData13_1") = NewData13_1.Text
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    If iPlanKind = 1 Then
                        sql = "Select * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "'"
                        dt = DbAccess.GetDataTable(sql, da, TransConn)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("ClassCount") = NewData13_1.Text
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da)
                        End If
                    End If
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-42C2")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                If iPlanKind = 1 Then
                    '2006/11/16 by Ellen update PlanInfo 
                    sql = "SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "'"
                    dt = DbAccess.GetDataTable(sql, da, TransConn)
                    If dt.Rows.Count <> 0 Then
                        dr = dt.Rows(0)
                        dr("ClassCount") = NewData13_1.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    End If
                End If

            Case Cst_i科場地 'As Integer=14 '學(術)科場地 '變更學(術)科場地
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    'Trans=DbAccess.BeginTrans(Transconn)
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i科場地, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)
                    '異動項目
                    'dr("OldData14_1")=If(OldData14_1.Value="", Convert.DBNull, OldData14_1.Value)
                    'dr("OldData14_2")=If(OldData14_2.Value="", Convert.DBNull, OldData14_2.Value)
                    dr("OldData14_1") = If(OldData14_1b.Value = "", Convert.DBNull, OldData14_1b.Value)
                    dr("OldData14_2") = If(OldData14_2b.Value = "", Convert.DBNull, OldData14_2b.Value)
                    dr("OldData14_3") = If(OldData14_3.Value = "", Convert.DBNull, OldData14_3.Value)
                    dr("OldData14_4") = If(OldData14_4.Value = "", Convert.DBNull, OldData14_4.Value)
                    'dr("NewData14_1")=If(NewData14_1.SelectedValue="0", Convert.DBNull, NewData14_1.SelectedValue)
                    'dr("NewData14_2")=If(NewData14_2.SelectedValue="0", Convert.DBNull, NewData14_2.SelectedValue)
                    dr("NewData14_1") = If(NewData14_1b.SelectedValue = "0", Convert.DBNull, NewData14_1b.SelectedValue)
                    dr("NewData14_2") = If(NewData14_2b.SelectedValue = "0", Convert.DBNull, NewData14_2b.SelectedValue)
                    dr("NewData14_3") = If(NewData14_3.SelectedValue = "0", Convert.DBNull, NewData14_3.SelectedValue)
                    dr("NewData14_4") = If(NewData14_4.SelectedValue = "0", Convert.DBNull, NewData14_4.SelectedValue)
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-A3D3")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i上課時間 'As Integer=15 '變更上課時間
                'Dim Trans As SqlTransaction=Nothing
                If Session("REVISE_ONCLASS") Is Nothing Then
                    Common.MessageBox(Page, "請輸入欲變更之上課時間!-C208")
                    Exit Sub
                End If
                If DataGrid2.Items.Count = 0 Then
                    Common.MessageBox(Page, "請輸入欲變更之上課時間!-2436")
                    Exit Sub
                End If

                Dim dtTemp As DataTable
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    'Trans=DbAccess.BeginTrans(Transconn)
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i上課時間, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    dtTemp = Session("REVISE_ONCLASS")
                    sql = " Select * FROM REVISE_ONCLASS WHERE 1<>1 "
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    For Each dr In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        'REVISE_ONCLASS_ROCID_SEQ
                        If dr("ROCID") < 0 Then dr("ROCID") = DbAccess.GetNewId(Trans, "REVISE_ONCLASS_ROCID_SEQ,REVISE_ONCLASS,ROCID")
                        dr("PlanID") = rPlanID 'Request("PlanID")
                        dr("ComIDNO") = rComIDNO 'Request("cid")
                        dr("SeqNO") = rSeqNo 'Request("no")
                        dr("SCDate") = ApplyDate.Text
                        dr("SubSeqNO") = iSubSeqNO
                    Next
                    dt = dtTemp.Copy
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-E1D8")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i其他 'As Integer=16 '其它
                Dim strErrmsg As String = ""
                OldData15_1.Text = TIMS.ClearSQM(OldData15_1.Text)
                NewData15_1.Text = TIMS.ClearSQM(NewData15_1.Text)
                'If OldData15_1.Text.Trim <> "" Then OldData15_1.Text=OldData15_1.Text.Trim
                'If NewData15_1.Text.Trim <> "" Then NewData15_1.Text=NewData15_1.Text.Trim
                If Len(OldData15_1.Text) > 1400 Then strErrmsg &= "原計畫內容資料過長(1400)有誤!" & vbCrLf
                If Len(NewData15_1.Text) > 1400 Then strErrmsg &= "變更內容資料過長(1400)有誤!" & vbCrLf
                If strErrmsg <> "" Then
                    Common.MessageBox(Page, strErrmsg)
                    Exit Sub
                End If
                '(存入申請異動項目)
                'Dim Trans As SqlTransaction=Nothing
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    'Trans=DbAccess.BeginTrans(Transconn)
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i其他, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    '異動項目
                    dr("OldData15_1") = If(OldData15_1.Text = "", Convert.DBNull, OldData15_1.Text)
                    dr("NewData15_1") = If(NewData15_1.Text = "", Convert.DBNull, NewData15_1.Text)
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-BCA2")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

            Case Cst_i報名日期 'As Integer=17  '20080825 andy  add 報名日期 變更報名起訖
                'Dim Time1, Time2 As String
                Dim Time1 As String = New_SEnterDate.Text + " " + String.Format("{0:00}", CType(HR1.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM1.SelectedValue, Integer))
                Dim Time2 As String = New_FEnterDate.Text + " " + String.Format("{0:00}", CType(HR2.SelectedValue, Integer)) + ":" + String.Format("{0:00}", CType(MM2.SelectedValue, Integer))
                ' start (檢查資料)
                If New_SEnterDate.Text = "" OrElse New_FEnterDate.Text = "" Then
                    Common.MessageBox(Page, "報名起訖日期欄位為空白！")
                    Exit Sub
                End If
                If CDate(Time1) > CDate(Time2) Then
                    Common.MessageBox(Page, "「起始時間」" + Time1 + "大於「結束時間」" + Time2 + "！")
                    Exit Sub
                ElseIf CDate(Time1) = CDate(Time2) Then
                    Common.MessageBox(Page, "「起始時間」" + Time1 + "不能與「結束時間」" + Time2 + "相同！")
                    Exit Sub
                End If

                '(存入申請異動項目)
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
                Try
                    sql = TIMS.GET_PlanRevise(Me, Get_NowDate(), Cst_i報名日期, Cst_sql, gobjconn)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    Call Utl_PlanReviseChkTableGetRow(dt, dr, v_ChgItem)

                    dr("OldData17_1") = If(Old_SEnterDate.Text = "", Convert.DBNull, Old_SEnterDate.Text)
                    dr("NewData17_1") = If(Time1 = "", Convert.DBNull, Time1)
                    dr("OldData17_2") = If(Old_FEnterDate.Text = "", Convert.DBNull, Old_FEnterDate.Text)
                    dr("NewData17_2") = If(Time2 = "", Convert.DBNull, Time2)
                    Dim htSS As New Hashtable From {
                        {"ReviseCont", ReviseCont.Text},
                        {"changeReason", v_changeReason}
                    }
                    Call INS_CMN1(dr, sm, iPlanKind, htSS)
                    DbAccess.UpdateDataTable(dt, da, Trans)
                    DbAccess.CommitTrans(Trans)
                Catch ex As Exception
                    Dim strErrmsg5 As String = ""
                    strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                    strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg5)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Common.MessageBox(Page, "計畫變更儲存失敗!-5A8D")
                    If flagDebugTest Then Throw ex
                    Exit Sub
                End Try

                Dim ocid As String = ""
                Dim drP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, gobjconn)
                If Not drP Is Nothing Then ocid = Convert.ToString(drP("ocid"))
                If iPlanKind = 1 AndAlso ocid <> "" Then
                    Dim Trans2 As SqlTransaction = DbAccess.BeginTrans(TransConn)
                    Try
                        sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ocid & "'"
                        dt = DbAccess.GetDataTable(sql, da, Trans2)
                        If dt.Rows.Count <> 0 Then
                            dr = dt.Rows(0)
                            dr("SEnterDate") = TIMS.Cdate2(Time1) '必填
                            dr("FEnterDate") = TIMS.Cdate2(Time2) '必填
                            dr("LastState") = "M" 'M: 修改(最後異動狀態)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now()
                            DbAccess.UpdateDataTable(dt, da, Trans2)
                        End If
                        DbAccess.CommitTrans(Trans2)
                    Catch ex As Exception
                        Dim strErrmsg5 As String = ""
                        strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                        strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                        strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                        strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                        Call TIMS.WriteTraceLog(strErrmsg5)

                        DbAccess.RollbackTrans(Trans2)
                        Call TIMS.CloseDbConn(TransConn)
                        Common.MessageBox(Page, "計畫變更儲存失敗!-26CC")
                        If flagDebugTest Then Throw ex
                        Exit Sub
                    End Try
                End If

        End Select

        Call TIMS.CloseDbConn(TransConn)

        If strMassage <> "" Then
            '檢查全日制規則,若是有錯誤訊息
            strMassage &= "計劃變更申請成功!\n"
            If iPlanKind = 1 Then strMassage &= "(自辦計畫 直接審核成功)\n"
            Dim sScript1 As String = ""
            sScript1 = ""
            sScript1 &= "<script language=javascript>alert('" + strMassage + "');"
            sScript1 &= "location.href='TC_05_001.aspx?ID=" & Request("ID") & "';</script>"
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(sScript1))
            Page.RegisterStartupScript("", sScript1)
        Else
            'Common.RespWrite(Me, "<script language=javascript>window.alert('計劃變更申請成功!');")
            'insert_next_val 0:(非)自辦計畫 /1:自辦計畫
            Dim strPjs1 As String = "var insert_next_val='0';" '計劃變更申請成功，是否繼續新增?
            If iPlanKind = 1 Then strPjs1 = "var insert_next_val='1';"
            ltlInserNextFlag.Text = "1" '啟動使用，注意html(aspx)不可有空白
            Common.AddClientScript(Page, strPjs1) '計劃變更申請成功，是否繼續新增?
        End If
    End Sub

    Private Sub EGenSci_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EGenSci.TextChanged
        Call SumESumSci()
    End Sub

    Private Sub EProSci_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EProSci.TextChanged
        Call SumESumSci()
    End Sub

    Private Sub SumESumSci()
        If EGenSci.Text = "" And EProSci.Text = "" Then
            ESumSci.Text = "0"
        ElseIf EGenSci.Text = "" Then
            ESumSci.Text = CStr(0 + CInt(EProSci.Text))
        ElseIf EProSci.Text = "" Then
            ESumSci.Text = CStr(0 + CInt(EGenSci.Text))
        Else
            ESumSci.Text = CStr(CInt(EProSci.Text) + CInt(EGenSci.Text))
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        'ShowCourseList(PlaceDate, SPlace, msg3)
        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Select Case v_ChgItem'.SelectedValue
            Case "3" '地點
                ShowCourseList(PlaceDate, SPlace, msg3, "room")
            Case Else
                ShowCourseList(PlaceDate, SPlace, msg3)
        End Select
    End Sub

    '師資變更 (日期選擇後顯示當天課程狀況。)
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If vsShowmsg4 = "" Then
            ShowCourseList(TechDate, STeacher, msg4, "Tech")
            vsShowmsg4 = "1"
        End If

        '師資設定
        OLessonTeah1.Style("display") = cst_inline1
        OLessonTeah2.Style("display") = cst_inline1
        OLessonTeah3.Style("display") = cst_inline1

        '當天無排課資料
        If Not STeacher.Visible Then
            OLessonTeah1.Style("display") = "none"
            OLessonTeah2.Style("display") = "none"
            OLessonTeah3.Style("display") = "none"
        Else
            '非自辦計畫使用
            If TIMS.Cst_TPlanID02Plan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                If hid_NoTechID2.Value = "Y" Then OLessonTeah2.Style("display") = "none"
                If hid_NoTechID3.Value = "Y" Then OLessonTeah3.Style("display") = "none"
            End If
        End If
    End Sub

    '變更前上課時間
    Private Sub PlanClassTime()
        'Dim sql As String=""
        Dim dt As DataTable = Nothing
        Select Case UCase(rActCheck)
            Case Cst_cRevise
                Dim sql_OLD As String = ""
                sql_OLD &= " Select ROCID,PLANID,COMIDNO,SEQNO,Weeks,Times FROM REVISE_ONCLASS_OLD"
                sql_OLD &= " WHERE PlanID =@PlanID And ComIDNO =@ComIDNO And SeqNO =@SeqNO And SCDate=@SCDate And SubSeqNO=@SubSeqNO"
                sql_OLD &= " ORDER BY ROCID"
                Dim parms As New Hashtable
                TIMS.SetMyValue2(parms, "PlanID", Convert.ToString(rPlanID)) '計畫PK
                TIMS.SetMyValue2(parms, "ComIDNO", rComIDNO) '計畫PK
                TIMS.SetMyValue2(parms, "SeqNO", Convert.ToString(rSeqNo)) '計畫PK
                TIMS.SetMyValue2(parms, "SCDate", If(rSCDate <> "", rSCDate, Convert.DBNull)) '計畫PK
                TIMS.SetMyValue2(parms, "SubSeqNO", iSubSeqNO) '計畫PK
                dt = DbAccess.GetDataTable(sql_OLD, gobjconn, parms)
                If dt.Rows.Count = 0 Then
                    Dim sql_R1 As String = ""
                    sql_R1 &= " Select POCID,PLANID,COMIDNO,SEQNO,Weeks,Times FROM PLAN_ONCLASS"
                    sql_R1 &= " WHERE PlanID =@PlanID And ComIDNO =@ComIDNO And SeqNO =@SeqNO"
                    sql_R1 &= " ORDER BY POCID"
                    Dim parms1 As New Hashtable
                    TIMS.SetMyValue2(parms1, "PlanID", Convert.ToString(rPlanID)) '計畫PK
                    TIMS.SetMyValue2(parms1, "ComIDNO", rComIDNO) '計畫PK
                    TIMS.SetMyValue2(parms1, "SeqNO", Convert.ToString(rSeqNo)) '計畫PK
                    dt = DbAccess.GetDataTable(sql_R1, gobjconn, parms1)
                End If
            Case Else
                Dim sql_N1 As String = ""
                sql_N1 &= " Select POCID,PLANID,COMIDNO,SEQNO,Weeks,Times FROM PLAN_ONCLASS"
                sql_N1 &= " WHERE PlanID =@PlanID And ComIDNO =@ComIDNO And SeqNO =@SeqNO"
                sql_N1 &= " ORDER BY POCID"
                Dim parms1 As New Hashtable
                TIMS.SetMyValue2(parms1, "PlanID", Convert.ToString(rPlanID)) '計畫PK
                TIMS.SetMyValue2(parms1, "ComIDNO", rComIDNO) '計畫PK
                TIMS.SetMyValue2(parms1, "SeqNO", Convert.ToString(rSeqNo)) '計畫PK
                dt = DbAccess.GetDataTable(sql_N1, gobjconn, parms1)
        End Select
        'dt.Columns("POCID").AutoIncrement=True
        'dt.Columns("POCID").AutoIncrementSeed=-1
        'dt.Columns("POCID").AutoIncrementStep=-1
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    '變更後上課時間table '若當日有申請，使用當日最後資料
    Function GET_REVISE_ONCLASS() As DataTable
        'parms.Clear()
        Dim parms As New Hashtable From {
            {"SCDate", TIMS.Cdate2(Get_NowDate())},
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNO", rSeqNo}
        }
        'sql &= "   And SCDate=" & TIMS.to_date(Now.Date.ToString("yyyy/MM/dd")) & vbCrLf
        'sql &= " And SCDate=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        Dim sql As String = ""
        sql &= " SELECT * FROM REVISE_ONCLASS" & vbCrLf
        sql &= " WHERE SCDate=@SCDate And PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO" & vbCrLf
        sql &= " ORDER BY SubSeqNO DESC" & vbCrLf 'MAX SubSeqNO
        Dim dt As DataTable = DbAccess.GetDataTable(sql, gobjconn, parms)
        If dt.Rows.Count = 0 Then Return dt '無資料直接返回

        Dim dr As DataRow = dt.Rows(0)
        iSubSeqNO = dr("SubSeqNO") 'MAX SubSeqNO

        Dim pms_2 As New Hashtable From {
            {"SCDate", TIMS.Cdate2(Get_NowDate())},
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNO", rSeqNo},
            {"SubSeqNO", iSubSeqNO}
        }
        'sql &= "   AND SCDate=" & TIMS.to_date(Now.Date.ToString("yyyy/MM/dd")) & vbCrLf
        'sql &= " AND SCDate=dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        Dim sql2 As String = ""
        sql2 &= " SELECT * FROM REVISE_ONCLASS" & vbCrLf
        sql2 &= " WHERE SCDate=@SCDate And PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO" & vbCrLf
        sql2 &= " AND SubSeqNO=@SubSeqNO" & vbCrLf
        dt = DbAccess.GetDataTable(sql2, gobjconn, pms_2) '有資料取當日最大值或異常
        Return dt
    End Function

    '變更後上課時間
    Private Sub ReviseClassTime(ByVal sType As String)
        'Optional ByVal sType As String=""
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        'sType 1:第1次呼叫 'Dim dr As DataRow 
        Select Case sType
            Case "1"
                'sql &= " SELECT * FROM REVISE_ONCLASS WHERE SCDate='" & Now.Date & "' and PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "'" & vbCrLf
                'sql &= " and SubSeqNO in (SELECT MAX(SubSeqNO) FROM REVISE_ONCLASS WHERE SCDate='" & Now.Date & "' and PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "')" & vbCrLf
                '變更後上課時間table '若當日有申請，使用當日最後資料
                dt = GET_REVISE_ONCLASS()
                dt.Columns("ROCID").AutoIncrement = True
                dt.Columns("ROCID").AutoIncrementSeed = -1
                dt.Columns("ROCID").AutoIncrementStep = -1
                DataGrid2.DataSource = dt
                DataGrid2.DataBind()

                Session("REVISE_ONCLASS") = dt
            Case Else
                If Session("REVISE_ONCLASS") Is Nothing Then
                    'sql="SELECT * FROM REVISE_ONCLASS WHERE SCDate='" & SCDate & "' and SubSeqNO=" & SubSeqNO & " and PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "'"
                    If rSCDate <> "" AndAlso iSubSeqNO <> 0 Then
                        sql = ""
                        sql &= " SELECT * FROM REVISE_ONCLASS "
                        sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                        sql &= " AND SCDate=" & TIMS.To_date(rSCDate) & " AND SubSeqNO =" & CStr(iSubSeqNO)
                    Else
                        sql = " SELECT * FROM REVISE_ONCLASS WHERE 1<>1 "
                    End If
                    dt = DbAccess.GetDataTable(sql, gobjconn)
                    dt.Columns("ROCID").AutoIncrement = True
                    dt.Columns("ROCID").AutoIncrementSeed = -1
                    dt.Columns("ROCID").AutoIncrementStep = -1
                    Session("REVISE_ONCLASS") = dt
                Else
                    dt = Session("REVISE_ONCLASS")
                End If

                Dim flag_CanUsefuncbutton As Boolean = True '(可使用功能鍵)
                If rActCheck = Cst_cRevise Then flag_CanUsefuncbutton = False '(停用功能鍵)
                If rActCheck = Cst_cRevise AndAlso rPARTREDUC1 = "Y" Then flag_CanUsefuncbutton = True '(可使用功能鍵)
                DataGrid2.Columns(2).Visible = If(flag_CanUsefuncbutton, True, False) '功能鍵狀態

                DataGrid2.DataSource = dt
                DataGrid2.DataBind()
                Session("REVISE_ONCLASS") = dt
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OldWeeks1 As Label = e.Item.FindControl("OldWeeks1")
                Dim OldTimes1 As Label = e.Item.FindControl("OldTimes1")
                OldWeeks1.Text = Convert.ToString(drv("Weeks"))
                OldTimes1.Text = If(Convert.ToString(drv("Times")).Length > cst_i_Times_c_max_length, Left(drv("Times").ToString, cst_i_Times_c_max_length), Convert.ToString(drv("Times")))
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        If e Is Nothing Then Return
        Select Case e.CommandName
            Case "edit"
                DataGrid2.EditItemIndex = e.Item.ItemIndex
            Case "del"
                Dim dt As DataTable = Session("REVISE_ONCLASS")
                If TIMS.dtNODATA(dt) Then Return

                If dt.Select("ROCID='" & e.CommandArgument & "'").Length <> 0 Then dt.Select("ROCID='" & e.CommandArgument & "'")(0).Delete()
                Session("REVISE_ONCLASS") = dt
                DataGrid2.DataSource = dt
                DataGrid2.EditItemIndex = -1
            Case "save"
                Dim dt As DataTable = Session("REVISE_ONCLASS")
                If TIMS.dtNODATA(dt) Then Return

                Dim Weeks As DropDownList = e.Item.FindControl("NewWeeks2")
                Dim Times As TextBox = e.Item.FindControl("NewTimes2")
                If dt.Select("ROCID='" & e.CommandArgument & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select("ROCID='" & e.CommandArgument & "'")(0)
                    dr("Weeks") = Weeks.SelectedValue
                    dr("Times") = If(Times.Text.Length > cst_i_Times_c_max_length, Left(Times.Text, cst_i_Times_c_max_length), Times.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                Session("REVISE_ONCLASS") = dt
                DataGrid2.EditItemIndex = -1
            Case "cancel"
                DataGrid2.EditItemIndex = -1
        End Select
        ReviseClassTime("")
        Page.RegisterStartupScript("Londing", "<script>window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim NewWeeks1 As Label = e.Item.FindControl("NewWeeks1")
                Dim NewTimes1 As Label = e.Item.FindControl("NewTimes1")
                Dim btn1 As Button = e.Item.FindControl("Button7")
                Dim btn2 As Button = e.Item.FindControl("Button8")
                btn1.Enabled = Button29.Enabled '同新增鈕
                btn2.Enabled = Button29.Enabled '同新增鈕
                NewWeeks1.Text = Convert.ToString(drv("Weeks")) '.ToString
                NewTimes1.Text = If(Convert.ToString(drv("Times")).Length > cst_i_Times_c_max_length, Left(drv("Times").ToString, cst_i_Times_c_max_length), Convert.ToString(drv("Times")))
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn2.CommandArgument = Convert.ToString(drv("ROCID"))

            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim NewWeeks2 As DropDownList = e.Item.FindControl("NewWeeks2")
                Dim NewTimes2 As TextBox = e.Item.FindControl("NewTimes2")
                Dim btn1 As Button = e.Item.FindControl("Button9")
                Dim btn2 As Button = e.Item.FindControl("Button10")

                NewWeeks2 = TIMS.Get_ddlWeeks(NewWeeks2)
                If Convert.ToString(drv("Weeks")) <> "" Then Common.SetListItem(NewWeeks2, Convert.ToString(drv("Weeks")))
                NewTimes2.Text = If(Convert.ToString(drv("Times")).Length > cst_i_Times_c_max_length, Left(drv("Times").ToString, cst_i_Times_c_max_length), Convert.ToString(drv("Times")))
                btn1.CommandArgument = Convert.ToString(drv("ROCID"))
        End Select
    End Sub

    '新增 上課時段
    Private Sub Button29_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim Errmsg As String = ""
        Dim v_Weeks As String = TIMS.GetListValue(Weeks)
        'Dim vsWeeks As String=TIMS.ClearSQM(Weeks.SelectedValue)
        txtTimes.Text = TIMS.ClearTimesFMT(txtTimes.Text)
        'Const cst_i_Times_c_max_length As Integer=200 'by AMU 20220310
        'Const cst_i_Times_c_min_length As Integer=5
        Dim i_Times_c_max_length As Integer = cst_i_Times_c_max_length
        Dim i_Times_c_min_length As Integer = cst_i_Times_c_min_length
        Dim s_err_msg1 As String = String.Format("上課時間／時間內容，長度超過限制範圍{0}文字長度", i_Times_c_max_length)
        Dim s_err_msg2 As String = String.Format("上課時間／時間內容，長度小於限制範圍{0}文字長度", i_Times_c_min_length)
        If txtTimes.Text <> "" Then
            If txtTimes.Text.ToString.Length > i_Times_c_max_length Then Errmsg &= s_err_msg1 & vbCrLf
            If txtTimes.Text.ToString.Length < i_Times_c_min_length Then Errmsg &= s_err_msg2 & vbCrLf
        Else
            Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
        End If
        'Dim i_Times_c_max_length As Integer=cst_i_Times_c_max_length
        'Dim i_Times_c_min_length As Integer=cst_i_Times_c_min_length
        'txtTimes.Text=TIMS.ClearSQM(txtTimes.Text)
        'If txtTimes.Text <> "" Then
        '    If txtTimes.Text.ToString.Length > i_Times_c_max_length Then Errmsg &= String.Format("上課時間／時間內容，長度超過限制範圍({0})文字長度", i_Times_c_max_length) & vbCrLf
        'Else
        '    Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
        'End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim Newdt As New DataTable
        Dim Newdr As DataRow = Nothing
        Dim sql As String = ""
        If Session("REVISE_ONCLASS") Is Nothing Then
            'sql="SELECT * FROM REVISE_ONCLASS WHERE SCDate='" & SCDate & "' and SubSeqNO=" & SubSeqNO & " and PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "'"
            If rSCDate <> "" AndAlso iSubSeqNO <> 0 Then
                sql = ""
                sql &= " SELECT * FROM REVISE_ONCLASS "
                sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                sql &= " AND SCDate=" & TIMS.To_date(rSCDate) & " AND SubSeqNO=" & CStr(iSubSeqNO)
            Else
                sql &= " SELECT * FROM REVISE_ONCLASS WHERE 1<>1 "
            End If
            Newdt = DbAccess.GetDataTable(sql, gobjconn)
            Newdt.Columns("ROCID").AutoIncrement = True
            Newdt.Columns("ROCID").AutoIncrementSeed = -1
            Newdt.Columns("ROCID").AutoIncrementStep = -1
        Else
            Newdt = Session("REVISE_ONCLASS")
        End If

        Newdr = Newdt.NewRow
        Newdt.Rows.Add(Newdr)
        Newdr("Weeks") = v_Weeks 'Weeks.SelectedValue()
        'Newdr("Times")=Times.Text
        'txtTimes.Text=TIMS.ClearTimesFMT(txtTimes.Text)
        If txtTimes.Text.Length > cst_i_Times_c_max_length Then txtTimes.Text = Left(txtTimes.Text, cst_i_Times_c_max_length)
        Newdr("Times") = If(txtTimes.Text <> "", txtTimes.Text, Convert.DBNull)
        Newdr("ModifyAcct") = sm.UserInfo.UserID
        Newdr("ModifyDate") = Now

        Session("REVISE_ONCLASS") = Newdt
        DataGrid2.DataSource = Newdt
        DataGrid2.DataBind()
        Page.RegisterStartupScript("Londing", "<script>window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    ''' <summary>
    ''' 一般計畫儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub But_Sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Sub.Click
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'TIMS
            But_Sub.Style("display") = "none" '一般計畫儲存
            But_Sub28.Style("display") = "none" '產投計畫確定
            But_Save28.Style("display") = "none" '產投計畫儲存
            'But_UPLOAD28.Style("display")="none" '產投計畫儲存(上傳檔案)
            Common.MessageBox(Me, String.Concat(TIMS.cst_NODATAMsg3, "-12FD"))
            Exit Sub
        End If
        If changeReason.SelectedIndex = 0 Then
            Common.MessageBox(Page, "請選擇變更原因!!")
            Exit Sub
        End If
        If ReviseCont.Text.Length > 250 Then
            Common.MessageBox(Page, "變更說明" + ReviseCont.Text.Length.ToString + "字" + "超過限制" + "250個字")
            Exit Sub
        End If
        OldData15_1.Text = TIMS.ClearSQM(OldData15_1.Text)
        NewData15_1.Text = TIMS.ClearSQM(NewData15_1.Text)
        If OldData15_1.Text.Length > 1400 Then
            Common.MessageBox(Page, "原計畫其他內容" + OldData15_1.Text.Length.ToString + "字" + "超過限制" + "1400個字")
            Exit Sub
        End If
        If NewData15_1.Text.Length > 1400 Then
            Common.MessageBox(Page, "變更後其他內容" + NewData15_1.Text.Length.ToString + "字" + "超過限制" + "1400個字")
            Exit Sub
        End If

        Call Save_Sub() '一般計畫儲存
    End Sub

    '20080630 andy 產生課程表  start --
    ''' <summary> 產生課程表 </summary>
    Sub CreateTrainDesc()
        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        'Dim sql_1, sql_2, sql, AltPTDRDataID, ReviseStatus As String
        Dim dt As DataTable = Nothing
        Dim dt2 As DataTable = Nothing
        'Dim TD_1 As HtmlTableCell=FindControl("TD_1")
        'Dim TD_2 As HtmlTableCell=FindControl("TD_2")
        'Dim dr As DataRow
        TD_1.Style("display") = cst_inline1
        TD_2.Style("display") = cst_inline1
        Dim iPTDRID As Integer = 0
        'If rActCheck=Cst_cRevise Then
        'End If
        Select Case rActCheck 'UCase(Request("check"))
            Case Cst_cRevise '變更結果
                'PLAN_REVISE
                Dim htSS As New Hashtable
                TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
                TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
                TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
                TIMS.SetMyValue2(htSS, "rCDate", rSCDate) 'Request("CDate")
                TIMS.SetMyValue2(htSS, "rSubNo", iSubSeqNO) 'Request("SubNo")
                Dim dr As DataRow = Get_PlanReviseDataRow(htSS, gobjconn)
                '查無傳入資訊
                If dr Is Nothing Then Exit Sub
                Dim v_ReviseStatus As String = "N"
                'ReviseStatus="N"
                If Convert.ToString(dr("ReviseStatus")) <> "" Then v_ReviseStatus = Convert.ToString(dr("ReviseStatus"))
                Dim AltPTDRDataID As String = Convert.ToString(dr("AltDataID"))
                'Dim dorplist As String
                Dim showlist As New ArrayList '產投用選項
                showlist.Add(Cst_i訓練期間)
                showlist.Add(Cst_i師資)
                showlist.Add(Cst_i助教)
                showlist.Add(Cst_i科場地)
                showlist.Add(Cst_i上課時間)
                showlist.Add(Cst_i課程表)
                showlist.Add(Cst_i遠距教學)
                Dim flag_showdg As Boolean = False '產投用選項
                For Each s_listVal As String In showlist
                    If AltPTDRDataID = s_listVal Then flag_showdg = True
                    If flag_showdg Then Exit For
                Next
                If Not flag_showdg Then
                    TD_1.Style("display") = "none"
                    TD_2.Style("display") = "none"
                    Exit Sub
                End If
        End Select

        '變更申請前
        Select Case rActCheck 'UCase(Request("check"))
            Case Cst_cPlan '"PLAN_PLANINFO" '申請
                'ChgItem.SelectedValue '儲存前的動作
                iPTDRID = Get_PTDRID(1, rPlanID, rComIDNO, rSeqNo, Today.ToString("yyyy/MM/dd"), ViewState(vs_SubSeqNO), v_ChgItem, gobjconn)
                '檢核資料狀況
                Dim flagRR As Boolean = Check_PlanTrainDescReviseItem(iPTDRID, TIMS.CINT1(v_ChgItem), gobjconn)
                If flagRR Then
                    dt = Get_PlanTrainDesc(rPlanID, rComIDNO, rSeqNo)
                    dt2 = Get_PlanTrainDescNewRevise(iPTDRID, v_ChgItem)
                Else
                    dt = Get_PlanTrainDesc(rPlanID, rComIDNO, rSeqNo)
                    dt2 = dt
                End If

            Case Cst_cRevise '"PLAN_REVISE"  '變更結果
                '傳進 'If Request("AltDataID") <> "" Then Common.SetListItem(ChgItem, Request("AltDataID"))
                iPTDRID = Get_PTDRID(2, rPlanID, rComIDNO, rSeqNo, Hid_rCDATE.Value, iSubSeqNO, rAltDataID, gobjconn)
                dt = Get_PlanTrainDescOldRevise(iPTDRID, rAltDataID, cst_now) '課程表申請變更前
                If dt.Rows.Count = 0 Then dt = Get_PlanTrainDescOldRevise(iPTDRID, rAltDataID, cst_old1) '課程表申請變更前
                dt2 = Get_PlanTrainDescNewRevise(iPTDRID, rAltDataID) '課程表申請變更後
        End Select
        dtlist11 = TIMS.Get_TechListPlanRdt1(rPlanID, rComIDNO, rSeqNo, gobjconn)
        dtlist20 = TIMS.Get_TechListPlanRdt2(rPlanID, rComIDNO, rSeqNo, gobjconn)

        'Dim rPlanID As String=TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        'Dim rComIDNO As String=TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        'Dim rSeqNo As String=TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        'Dim SCDate As String=TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        'Dim SubSeqNo As String=TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO

        Select Case v_ChgItem'.SelectedValue '20
            Case Cst_i師資
                '產投用選項 'NewData11_1
                Dim htSS As New Hashtable
                TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
                TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
                TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
                TIMS.SetMyValue2(htSS, "SCDate", rSCDate) 'Request("CDate")
                TIMS.SetMyValue2(htSS, "SubSeqNo", iSubSeqNO) 'Request("SubNo")
                TIMS.SetMyValue2(htSS, "ActCheck", rActCheck) 'rActCheck / Cst_cPlan '申請 /Cst_cRevise '審核查詢
                htSS.Add("RID", RIDValue.Value)
                htSS.Add("TECHIDs", NewData11_1.Value)
                htSS.Add("TechTYPE", "A")
                SHOW_REVISE_TEACHER12(htSS, gobjconn)

            Case Cst_i助教
                '產投用選項 'NewData20_1
                Dim htSS As New Hashtable
                TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
                TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
                TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
                TIMS.SetMyValue2(htSS, "SCDate", rSCDate) 'Request("CDate")
                TIMS.SetMyValue2(htSS, "SubSeqNo", iSubSeqNO) 'Request("SubNo")
                TIMS.SetMyValue2(htSS, "ActCheck", rActCheck) 'rActCheck / Cst_cPlan '申請 /Cst_cRevise '審核查詢
                htSS.Add("RID", RIDValue.Value)
                htSS.Add("TECHIDs", NewData20_1.Value)
                htSS.Add("TechTYPE", "B")
                SHOW_REVISE_TEACHER12(htSS, gobjconn)

            Case Else '課程表:18

        End Select

        If dt Is Nothing Then
            ViewState("Reflash_Plan_TrainDesc") = "N"
            Datagrid3Table.Visible = False
            Datagrid3Table.Style.Item("display") = "none"

            ViewState("PTDRID") = "0"
            ViewState("AltDataID") = ""
        Else
            Datagrid3Table.Style.Item("display") = cst_inline1

            Dim fg_show_EHour As Boolean = If(hid_TMID.Value = TIMS.cst_EHour_Use_TMID, True, False)
            DataGrid4.Columns(cst_DG4_EHour_技檢訓練時數_iCOL).Visible = fg_show_EHour
            DataGrid4.DataSource = dt '課程表申請變更前
            DataGrid4.DataBind()

            DataGrid3.Columns(cst_DG3_EHour_技檢訓練時數_iCOL).Visible = fg_show_EHour
            DataGrid3.DataSource = dt2 '課程表申請變更後
            DataGrid3.DataBind()

            ViewState("PTDRID") = iPTDRID
            ViewState("AltDataID") = v_ChgItem 'ChgItem.SelectedValue
        End If
    End Sub

    '課程表申請變更後
    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim HIC_Choose1 As HtmlInputCheckBox = e.Item.FindControl("Choose1")
                If HIC_Choose1 Is Nothing Then Return
                HIC_Choose1.Disabled = True
                HIC_Choose1.Style("display") = "none"
                If (rActCheck.Equals(Cst_cPlan)) Then 'PLAN_PLANINFO '申請中
                    Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
                    Select Case v_ChgItem'.SelectedValue
                        Case Cst_i遠距教學
                            HIC_Choose1.Style("display") = cst_inline1
                            HIC_Choose1.Disabled = False
                    End Select
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drData As DataRowView = e.Item.DataItem
                Dim hidePTDRID As HtmlInputHidden = e.Item.FindControl("hide_PTDRID")
                Dim hidePTDID As HtmlInputHidden = e.Item.FindControl("hide_PTDID")
                Dim hideID1 As HtmlInputHidden = e.Item.FindControl("hide_ID1")
                Dim hideID2 As HtmlInputHidden = e.Item.FindControl("hide_ID2")
                Dim hideID3 As HtmlInputHidden = e.Item.FindControl("hide_ID3") '時數
                Dim hideID9 As HtmlInputHidden = e.Item.FindControl("hide_ID9") '技檢訓練時數
                Dim hideID4 As HtmlInputHidden = e.Item.FindControl("hide_ID4")
                Dim txtSTrainDate As TextBox = e.Item.FindControl("txt_STrainDate")
                Dim hideSTrainDate As HtmlInputHidden = e.Item.FindControl("hide_STrainDate")
                Dim txtPName As TextBox = e.Item.FindControl("txt_PName")
                Dim hidePName As HtmlInputHidden = e.Item.FindControl("hide_PName")
                Dim txtPHour As TextBox = e.Item.FindControl("txt_PHour") '時數
                Dim hidePHour As HtmlInputHidden = e.Item.FindControl("hide_PHour") '時數
                Dim txtEHour As TextBox = e.Item.FindControl("txt_EHour") '技檢訓練時數
                Dim hideEHour As HtmlInputHidden = e.Item.FindControl("hide_EHour") '技檢訓練時數
                Dim txtPCont As TextBox = e.Item.FindControl("txt_PCont")
                Dim listClassification As DropDownList = e.Item.FindControl("list_Classification")
                Dim listPTID As DropDownList = e.Item.FindControl("list_PTID") '上課地點
                Dim hidePTID As HtmlInputHidden = e.Item.FindControl("hide_PTID")
                Dim hideID5 As HtmlInputHidden = e.Item.FindControl("hide_ID5")
                Dim listTechID As DropDownList = e.Item.FindControl("list_TechID")
                Dim hideTechID As HtmlInputHidden = e.Item.FindControl("hide_TechID")
                Dim hide_ID6 As HtmlInputHidden = e.Item.FindControl("hide_ID6")
                Dim listTechID2 As DropDownList = e.Item.FindControl("list_TechID2")
                Dim hideTechID2 As HtmlInputHidden = e.Item.FindControl("hide_TechID2")

                Dim TPERIOD28_1t As CheckBox = e.Item.FindControl("TPERIOD28_1t")
                Dim TPERIOD28_2t As CheckBox = e.Item.FindControl("TPERIOD28_2t")
                Dim TPERIOD28_3t As CheckBox = e.Item.FindControl("TPERIOD28_3t")
                Dim hidTPERIOD28 As HtmlInputHidden = e.Item.FindControl("hidTPERIOD28")
                Dim hide_ID7 As HtmlInputHidden = e.Item.FindControl("hide_ID7")
                'Cst_i遠距教學
                Dim bx_FARLEARN As CheckBox = e.Item.FindControl("bx_FARLEARN")
                Dim hide_FARLEARN As HtmlInputHidden = e.Item.FindControl("hide_FARLEARN")
                Dim hide_ID8 As HtmlInputHidden = e.Item.FindControl("hide_ID8")
                bx_FARLEARN.Attributes.Add("onClick", "reset_Choose1();")
                bx_FARLEARN.Checked = If(Convert.ToString(drData("FARLEARN")).Equals("Y"), True, False)
                bx_FARLEARN.Enabled = False
                hide_FARLEARN.Value = Convert.ToString(drData("FARLEARN"))
                hide_ID8.Value = Convert.ToString(drData("ID8"))

                Dim img9 As HtmlImage = e.Item.FindControl("Img9")
                Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)

                txtSTrainDate.Enabled = False
                img9.Visible = False '無法選擇日期
                txtPName.Enabled = False '授課時間
                txtPHour.Enabled = False '授課時間-時數
                txtEHour.Enabled = False '授課時間-技檢訓練時數
                TIMS.Tooltip(txtEHour, cst_EHour_t1, True)
                TPERIOD28_1t.Enabled = False '授課時間-授課時段-1
                TPERIOD28_2t.Enabled = False '授課時間-授課時段-2
                TPERIOD28_3t.Enabled = False '授課時間-授課時段-3
                listPTID.Enabled = False
                listTechID.Enabled = False
                listTechID2.Enabled = False
                listClassification.Enabled = False
                hidePTDRID.Value = Convert.ToString(drData("PTDRID"))
                hidePTDID.Value = Convert.ToString(drData("PTDID"))
                '日期
                hideID1.Value = Convert.ToString(drData("ID1"))
                txtSTrainDate.Text = Convert.ToString(drData("STrainDate"))
                hideSTrainDate.Value = TIMS.Cdate3(Convert.ToString(drData("STrainDate")))
                '授課時間
                hideID2.Value = Convert.ToString(drData("ID2"))
                txtPName.Text = Convert.ToString(drData("PName"))
                hidePName.Value = Convert.ToString(drData("PName"))
                '授課時間-時數
                hideID3.Value = Convert.ToString(drData("ID3"))
                txtPHour.Text = Convert.ToString(drData("PHour"))
                hidePHour.Value = Convert.ToString(drData("PHour"))
                '授課時間-技檢訓練時數
                hideID9.Value = Convert.ToString(drData("ID9"))
                txtEHour.Text = Convert.ToString(drData("EHour"))
                hideEHour.Value = Convert.ToString(drData("EHour"))
                '授課時間-授課時段
                Dim str_v_TPERIOD28 As String = If(Convert.ToString(drData("TPERIOD28")).Length >= 3, Convert.ToString(drData("TPERIOD28")), cst_NNN)
                TPERIOD28_1t.Checked = If(str_v_TPERIOD28.Substring(0, 1) = "Y", True, False)
                TPERIOD28_2t.Checked = If(str_v_TPERIOD28.Substring(1, 1) = "Y", True, False)
                TPERIOD28_3t.Checked = If(str_v_TPERIOD28.Substring(2, 1) = "Y", True, False)
                hidTPERIOD28.Value = str_v_TPERIOD28 'Convert.ToString(drData("TPERIOD28"))
                hide_ID7.Value = Convert.ToString(drData("ID7"))

                '課程進度/內容
                txtPCont.Text = TIMS.ClearSQM(Convert.ToString(drData("PCont")))
                '學/術科
                If Convert.ToString(drData("Classification1")) <> "" Then
                    Common.SetListItem(listClassification, Convert.ToString(drData("Classification1")))
                    'listClassification.SelectedValue=Convert.ToString(drData("Classification1"))
                End If
                '上課地點
                Dim newPTID As String = String.Empty
                Dim i_newPTID1 As Integer = 0
                Dim i_newPTID2 As Integer = 0
                hideID4.Value = Convert.ToString(drData("ID4"))
                '師資'NewData11_1
                hideID5.Value = Convert.ToString(drData("ID5"))
                Select Case TIMS.CINT1(v_ChgItem)'.SelectedValue '11
                    Case Cst_i師資
                        listTechID = TIMS.Get_TechListPlanR2(listTechID, NewData11_1.Value, dtlist11, gobjconn)
                    Case Else '課程表:18
                        listTechID = TIMS.Get_TechListPlanR2(listTechID, "", dtlist11, gobjconn)
                End Select
                If listTechID.Items.IndexOf(listTechID.Items.FindByValue(Convert.ToString(drData("TechID")))) <> -1 Then
                    If Convert.ToString(drData("TechID")) <> "" Then
                        Common.SetListItem(listTechID, Convert.ToString(drData("TechID")))
                        'listTechID.SelectedValue=Convert.ToString(drData("TechID"))
                    End If
                Else
                    listTechID.ClearSelection()
                End If
                hideTechID.Value = Convert.ToString(drData("TechID"))

                '助教'NewData20_1
                hide_ID6.Value = Convert.ToString(drData("ID6"))
                Select Case TIMS.CINT1(v_ChgItem)'.SelectedValue '20
                    Case Cst_i助教
                        '產投用選項
                        listTechID2 = TIMS.Get_TechListPlanR2(listTechID2, NewData20_1.Value, dtlist20, gobjconn)
                    Case Else '課程表:18
                        listTechID2 = TIMS.Get_TechListPlanR2(listTechID2, "", dtlist20, gobjconn)
                End Select
                If listTechID2.Items.IndexOf(listTechID2.Items.FindByValue(Convert.ToString(drData("TechID2")))) <> -1 Then
                    If Convert.ToString(drData("TechID2")) <> "" Then Common.SetListItem(listTechID2, Convert.ToString(drData("TechID2")))
                Else
                    listTechID2.ClearSelection()
                End If
                hideTechID2.Value = Convert.ToString(drData("TechID2"))

                Dim flag_can_clear As Boolean = False '清理不顯示的上課地點
                Dim flag_can_EDIT_1 As Boolean = False 'TRUE:目前可以編輯  FALSE:不可編輯
                Select Case rActCheck 'Request("check")
                    Case Cst_cPlan 'PLAN_PLANINFO '申請中
                        flag_can_EDIT_1 = True '目前可以編輯

                    Case Cst_cRevise 'PLAN_REVISE
                        flag_can_clear = If(rPARTREDUC1 = "Y", False, True) 'False:已送審(目前還原中) ／'True: 已送審了啦 
                        If rPARTREDUC1 = "Y" Then flag_can_EDIT_1 = True '目前可以編輯 '已送審(目前還原中)
                        If rPARTREDUC1 <> "Y" Then  '已送審(目前還原中) '查看
                            listPTID = Get_ListPTID(listPTID, drData("PTID"), drData("PTID"), rComIDNO, 2)
                            newPTID = drData("PTID")
                        End If
                End Select

                If flag_can_EDIT_1 Then
                    'Select Case v_ChgItem'.SelectedValue
                    '    Case Cst_i師資, Cst_i助教
                    'End Select
                    Select Case Convert.ToString(drData("Classification1"))
                        Case "1" '學科地點
                            '上課地點
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "cid", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, gobjconn)
                            listPTID = TIMS.Get_SciPTID(listPTID, Hid_ComIDNO.Value, 3, gobjconn)
                            newPTID = drData("PTID")
                        Case "2" '術科地點
                            '上課地點
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "cid", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, gobjconn)
                            listPTID = TIMS.Get_TechPTID(listPTID, Hid_ComIDNO.Value, 3, gobjconn)
                            newPTID = drData("PTID")
                    End Select
                    '申請中 '產投用選項
                    Select Case TIMS.CINT1(v_ChgItem)'.SelectedValue
                        Case Cst_i訓練期間
                            txtSTrainDate.Enabled = True
                            img9.Visible = True
                            If Convert.ToString(drData("STrainDate")) = "" Then
                                txtSTrainDate.Text = "請選擇日期"
                            Else
                                If CDate(drData("STrainDate")) < CDate(ASDate.Text) OrElse CDate(drData("STrainDate")) > CDate(AEDate.Text) Then txtSTrainDate.Text = "請選擇日期"
                            End If
                            Dim vvSTrainDate As String = If(IsDate(txtSTrainDate.Text), Common.FormatDate(txtSTrainDate.Text), "")
                            If (ASDate.Text = "") Then ASDate.Text = STDate.Value
                            If (AEDate.Text = "") Then AEDate.Text = FDDate.Value
                            img9.Attributes.Add("onClick", "openCalendar('" & txtSTrainDate.ClientID & "','" & ASDate.Text & "','" & AEDate.Text & "','" & vvSTrainDate & "');")
                        Case Cst_i師資
                            listTechID.Enabled = True
                        Case Cst_i助教
                            listTechID2.Enabled = True
                        Case Cst_i科場地 '上課地點
                            listPTID.Enabled = True
                            'Dim f_CanUpdate As Boolean=(Hid_PARTREDUC_Y_CanUpdate.Value="Y" AndAlso Convert.ToString(drData("PTID")) <> "" AndAlso TIMS.IsNumeric1(drData("PTID")))
                            If Convert.ToString(drData("PTID")) <> "" AndAlso TIMS.IsNumeric1(drData("PTID")) Then
                                Select Case Convert.ToString(drData("Classification1"))
                                    Case "1"
                                        listPTID = Get_ListPTID(listPTID, NewData14_1b.SelectedValue, NewData14_3.SelectedValue, rComIDNO, 1)
                                        i_newPTID1 = Get_TrainPlacePTID(NewData14_1b.SelectedValue, rComIDNO)
                                        i_newPTID2 = Get_TrainPlacePTID(NewData14_3.SelectedValue, rComIDNO)
                                    Case "2"
                                        listPTID = Get_ListPTID(listPTID, NewData14_2b.SelectedValue, NewData14_4.SelectedValue, rComIDNO, 1)
                                        i_newPTID1 = Get_TrainPlacePTID(NewData14_2b.SelectedValue, rComIDNO)
                                        i_newPTID2 = Get_TrainPlacePTID(NewData14_4.SelectedValue, rComIDNO)
                                End Select
                                newPTID = CStr(If(i_newPTID1 = Val(drData("PTID")) OrElse i_newPTID2 = Val(drData("PTID")), drData("PTID"), i_newPTID1))
                            Else
                                Select Case Convert.ToString(drData("Classification1"))
                                    Case "1"
                                        listPTID = Get_ListPTID(listPTID, NewData14_1b.SelectedValue, NewData14_3.SelectedValue, rComIDNO, 1)
                                        newPTID = Get_TrainPlacePTID(NewData14_1b.SelectedValue, rComIDNO)
                                    Case "2"
                                        listPTID = Get_ListPTID(listPTID, NewData14_2b.SelectedValue, NewData14_4.SelectedValue, rComIDNO, 1)
                                        newPTID = Get_TrainPlacePTID(NewData14_2b.SelectedValue, rComIDNO)
                                End Select
                            End If
                        Case Cst_i上課時間
                            txtPName.Enabled = True '授課時間
                            txtPHour.Enabled = True '授課時間-時數
                            txtEHour.Enabled = True '授課時間-'技檢訓練時數
                            TPERIOD28_1t.Enabled = True '授課時間-授課時段-1
                            TPERIOD28_2t.Enabled = True '授課時間-授課時段-2
                            TPERIOD28_3t.Enabled = True '授課時間-授課時段-3
                        Case Cst_i遠距教學
                            bx_FARLEARN.Enabled = True '遠距教學
                        Case Cst_i課程表
                            'PLAN_TRAINPLACE
                            txtSTrainDate.Enabled = True '日期
                            img9.Visible = True '日期選擇鈕
                            txtPName.Enabled = True '授課時間
                            txtPHour.Enabled = True '授課時間-時數
                            txtEHour.Enabled = True '授課時間-技檢訓練時數
                            TPERIOD28_1t.Enabled = True '授課時間-授課時段-1
                            TPERIOD28_2t.Enabled = True '授課時間-授課時段-2
                            TPERIOD28_3t.Enabled = True '授課時間-授課時段-3
                            'txtPCont.Enabled=True '不可變更
                            'listClassification.Enabled=True '不可變更
                            listPTID.Enabled = True
                            listTechID.Enabled = True '產投任課教師
                            listTechID2.Enabled = True '不可變更(開放)'產投助教
                            Select Case Convert.ToString(drData("Classification1"))
                                Case "1"
                                    listPTID = Get_ListPTID(listPTID, Convert.ToString(drData("SciPlaceID")), Convert.ToString(drData("SciPlaceID2")), rComIDNO, 1)
                                    newPTID = Convert.ToString(drData("PTID"))
                                Case "2" '產投助教
                                    listPTID = Get_ListPTID(listPTID, Convert.ToString(drData("TechPlaceID")), Convert.ToString(drData("TechPlaceID2")), rComIDNO, 1)
                                    newPTID = Convert.ToString(drData("PTID"))
                            End Select
                            Dim vvSTrainDate As String = If(IsDate(txtSTrainDate.Text), Common.FormatDate(txtSTrainDate.Text), "")
                            If (ASDate.Text = "") Then ASDate.Text = STDate.Value
                            If (AEDate.Text = "") Then AEDate.Text = FDDate.Value
                            img9.Attributes.Add("onClick", "openCalendar('" & txtSTrainDate.ClientID & "','" & ASDate.Text & "','" & AEDate.Text & "','" & vvSTrainDate & "');")
                    End Select
                End If

                'SetListItem
                If listPTID.Items.IndexOf(listPTID.Items.FindByValue(newPTID)) <> -1 Then
                    Common.SetListItem(listPTID, newPTID)
                    '整理下拉，只留選擇值
                    If (flag_can_clear) Then TIMS.GET_NewListItemVal(listPTID, newPTID)
                    'listPTID.SelectedValue=newPTID
                Else
                    If listPTID.Items.IndexOf(listPTID.Items.FindByValue(Convert.ToString(drData("PTID")))) <> -1 Then
                        Common.SetListItem(listPTID, Convert.ToString(drData("PTID")))
                        '整理下拉，只留選擇值
                        If (flag_can_clear) Then TIMS.GET_NewListItemVal(listPTID, Convert.ToString(drData("PTID")))
                        'listPTID.SelectedValue=Convert.ToString(drData("PTID"))
                    Else
                        listPTID.ClearSelection()
                    End If
                End If
                hidePTID.Value = Convert.ToString(drData("PTID"))
        End Select
    End Sub

    '課程表修改前(課程表申請變更前)
    Private Sub DataGrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'Show
                Dim drv As DataRowView = e.Item.DataItem
                Dim OldSTrainDateLabel As Label = e.Item.FindControl("OldSTrainDateLabel")
                Dim OldPNameLabel As Label = e.Item.FindControl("OldPNameLabel")
                Dim OldPHourLabel As Label = e.Item.FindControl("OldPHourLabel") '時數
                Dim OldEHourLabel As Label = e.Item.FindControl("OldEHourLabel") '技檢訓練時數
                Dim OldTPERIOD28_1t As CheckBox = e.Item.FindControl("OldTPERIOD28_1t")
                Dim OldTPERIOD28_2t As CheckBox = e.Item.FindControl("OldTPERIOD28_2t")
                Dim OldTPERIOD28_3t As CheckBox = e.Item.FindControl("OldTPERIOD28_3t")

                Dim OldPContText As TextBox = e.Item.FindControl("OldPContText")
                Dim OlddrpClassification1 As DropDownList = e.Item.FindControl("OlddrpClassification1")
                Dim OlddrpPTID As DropDownList = e.Item.FindControl("OlddrpPTID")
                Dim OldTech1Value As HtmlInputHidden = e.Item.FindControl("OldTech1Value")
                Dim OldTech1Text As TextBox = e.Item.FindControl("OldTech1Text")
                Dim OldTech2Value As HtmlInputHidden = e.Item.FindControl("OldTech2Value")
                Dim OldTech2Text As TextBox = e.Item.FindControl("OldTech2Text")
                'Cst_i遠距教學
                Dim OldFARLEARN As CheckBox = e.Item.FindControl("OldFARLEARN")
                OldFARLEARN.Checked = If(Convert.ToString(drv("FARLEARN")).Equals("Y"), True, False)
                'Dim btn19 As Button=e.Item.FindControl("Button19")
                'Dim btn20 As Button=e.Item.FindControl("Button20")
                'Dim tb_aid2 As TextBox=e.Item.FindControl("tb_aid2")
                'tb_aid2.Text=Convert.ToString(e.Item.ItemIndex + 1)
                'btn20.Attributes("onclick")=TIMS.cst_confirm_delmsg1
                'btn20.CommandArgument=drv("PTDID").ToString
                If Convert.ToString(drv("STrainDate")) <> "" Then OldSTrainDateLabel.Text = TIMS.Cdate3(drv("STrainDate"))
                OldPNameLabel.Text = drv("PName").ToString
                OldPHourLabel.Text = Convert.ToString(drv("PHour")) '時數
                OldEHourLabel.Text = Convert.ToString(drv("EHour")) '技檢訓練時數
                '授課時間-授課時段
                Dim str_v_TPERIOD28 As String = If(Convert.ToString(drv("TPERIOD28")).Length >= 3, Convert.ToString(drv("TPERIOD28")), cst_NNN)
                OldTPERIOD28_1t.Checked = If(str_v_TPERIOD28.Substring(0, 1) = "Y", True, False)
                OldTPERIOD28_2t.Checked = If(str_v_TPERIOD28.Substring(1, 1) = "Y", True, False)
                OldTPERIOD28_3t.Checked = If(str_v_TPERIOD28.Substring(2, 1) = "Y", True, False)

                OldPContText.Text = TIMS.ClearSQM(Convert.ToString(drv("PCont")))
                If drv("Classification1").ToString <> "" Then
                    Common.SetListItem(OlddrpClassification1, drv("Classification1").ToString)
                    Select Case OlddrpClassification1.SelectedValue
                        Case "1" '學科
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "cid", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, gobjconn)
                            OlddrpPTID = TIMS.Get_SciPTID(OlddrpPTID, Hid_ComIDNO.Value, 3, gobjconn)
                        Case "2" '術科
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "cid", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, gobjconn)
                            OlddrpPTID = TIMS.Get_TechPTID(OlddrpPTID, Hid_ComIDNO.Value, 3, gobjconn)
                    End Select
                    If Convert.ToString(drv("PTID")) <> "" AndAlso OlddrpPTID.SelectedIndex <> -1 Then
                        Common.SetListItem(OlddrpPTID, Convert.ToString(drv("PTID")))
                        '整理下拉，只留選擇
                        TIMS.GET_NewListItemVal(OlddrpPTID, Convert.ToString(drv("PTID")))
                    End If
                End If
                If drv("TechID").ToString <> "" Then
                    OldTech1Value.Value = drv("TechID").ToString
                    OldTech1Text.Text = TIMS.Get_TeachCName(drv("TechID"), gobjconn) 'TIMS.Get_TeacherName(drv("TechID").ToString)
                End If
                If drv("TechID2").ToString <> "" Then
                    OldTech2Value.Value = drv("TechID2").ToString
                    OldTech2Text.Text = TIMS.Get_TeachCName(drv("TechID2"), gobjconn) 'TIMS.Get_TeacherName(drv("TechID2").ToString)
                End If
        End Select
        'If (e.Item.ItemType=ListItemType.EditItem) Then
        '    Dim i As Integer
        '    For i=0 To e.Item.Cells.Count - 1
        '        e.Item.Cells(i).BackColor=System.Drawing.Color.FromName("#ffccff")
        '    Next
        'End If
    End Sub

    ''' <summary>
    ''' 儲存變更課程表-產投 '(若有錯誤，則會跳出錯誤文字)
    ''' </summary>
    ''' <param name="ApplyItem"></param>
    ''' <returns></returns>
    Function CHK_SAVE_TrainDescRserve(ByVal ApplyItem As Integer) As String
        'Dim DateNow As String=DateTime.Now.ToString("yyyy-MM-dd HH:mm@ss")
        Dim rst As String = "" '有值異常
        Dim sqlAdp As New SqlDataAdapter

        Dim errMsg As String = String.Empty

        Dim pms_CHK_SAVE_TrainDescRserve As New Hashtable
        Dim iPTDRID As Integer = ViewState("PTDRID")
        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Dim vrbl_DISTANCE As String = TIMS.GetListValue(rbl_DISTANCE)
        'DataGrid3
        Dim s_DISTANCE_N As String = TIMS.GET_DISTANCE_N(0, vrbl_DISTANCE)
        Dim i_DG3ROWS As Integer = 0
        Dim i_FARLEARN As Integer = 0
        For Each eItem As DataGridItem In DataGrid3.Items
            Dim hidePTDRID As HtmlInputHidden = eItem.FindControl("hide_PTDRID")
            Dim hidePTDID As HtmlInputHidden = eItem.FindControl("hide_PTDID")
            'Dim hideID1 As HtmlInputHidden=eItem.FindControl("hide_ID1")
            'Dim hideID2 As HtmlInputHidden=eItem.FindControl("hide_ID2")
            'Dim hideID3 As HtmlInputHidden=eItem.FindControl("hide_ID3")
            'Dim hideID4 As HtmlInputHidden=eItem.FindControl("hide_ID4")
            'Dim hideID5 As HtmlInputHidden=eItem.FindControl("hide_ID5")
            Dim txtSTrainDate As TextBox = eItem.FindControl("txt_STrainDate")
            Dim txtPName As TextBox = eItem.FindControl("txt_PName")
            Dim txtPHour As TextBox = eItem.FindControl("txt_PHour") '時數
            Dim txtEHour As TextBox = eItem.FindControl("txt_EHour") '技檢訓練時數
            Dim listPTID As DropDownList = eItem.FindControl("list_PTID")
            Dim listTechID As DropDownList = eItem.FindControl("list_TechID")
            Dim listTechID2 As DropDownList = eItem.FindControl("list_TechID2")
            'Dim hide_ID7 As HtmlInputHidden=eItem.FindControl("hide_ID7")
            Dim TPERIOD28_1t As CheckBox = eItem.FindControl("TPERIOD28_1t")
            Dim TPERIOD28_2t As CheckBox = eItem.FindControl("TPERIOD28_2t")
            Dim TPERIOD28_3t As CheckBox = eItem.FindControl("TPERIOD28_3t")
            Dim hidTPERIOD28 As HtmlInputHidden = eItem.FindControl("hidTPERIOD28")
            'Cst_i遠距教學
            Dim bx_FARLEARN As CheckBox = eItem.FindControl("bx_FARLEARN")
            'Dim hide_FARLEARN As HtmlInputHidden=eItem.FindControl("hide_FARLEARN")
            'Dim hide_ID8 As HtmlInputHidden=eItem.FindControl("hide_ID8")
            'Dim sFARLEARN As String=If(bx_FARLEARN.Checked, "Y", "")
            Select Case v_ChgItem'.SelectedValue
                Case Cst_i訓練期間
                    If txtSTrainDate.Text <> "" AndAlso IsDate(txtSTrainDate.Text) Then
                        txtSTrainDate.Style("background-color") = "#FFFFFF"
                        If CDate(txtSTrainDate.Text) < CDate(ASDate.Text) Or CDate(txtSTrainDate.Text) > CDate(AEDate.Text) Then
                            errMsg &= "日期未完全輸入 或 輸入區間錯誤(超出訓練期間範圍)。" & vbCrLf
                            txtSTrainDate.Style("background-color") = "#FFCCFF"
                        End If
                    Else
                        errMsg &= "日期未完全輸入 或 輸入區間錯誤(超出訓練期間範圍)。" & vbCrLf
                        txtSTrainDate.Style("background-color") = "#FFCCFF"
                    End If

                Case Cst_i師資
                    listTechID.Style("background-color") = "#FFFFFF"
                    If listTechID.SelectedIndex = 0 Then
                        errMsg &= "尚有任課教師未選擇。" & vbCrLf
                        listTechID.Style("background-color") = "#FFCCFF"
                    End If

                    txtPHour.Text = TIMS.ClearSQM(txtPHour.Text)
                    Dim v_listTechID As String = TIMS.GetListValue(listTechID)
                    'Dim v_listTechID2 As String=TIMS.GetListValue(listTechID2)
                    If v_listTechID <> "" AndAlso TIMS.VAL1(txtPHour.Text) > 0 Then
                        If CHK_TechID_PHOUR_Exceed_54_Hour(pms_CHK_SAVE_TrainDescRserve, String.Concat("T1x", v_listTechID), TIMS.VAL1(txtPHour.Text)) Then
                            Dim txt_listTechID As String = TIMS.GetListText(listTechID)
                            Dim s_ERR1 As String = String.Concat(txt_listTechID, "-", "授課時數不得超過54小時。(", v_listTechID, ")", vbCrLf)
                            If errMsg.IndexOf(s_ERR1) = -1 Then errMsg &= s_ERR1
                        End If
                    End If

                Case Cst_i助教 '產投用選項
                    listTechID2.Style("background-color") = "#FFFFFF"
                    If listTechID2.SelectedIndex = 0 Then
                        'errMsg &= "尚有助教未選擇。" & vbCrLf
                        listTechID2.Style("background-color") = "#FFCCFF"
                    End If

                    txtPHour.Text = TIMS.ClearSQM(txtPHour.Text)
                    ''Dim v_listTechID As String=TIMS.GetListValue(listTechID)
                    'Dim v_listTechID2 As String=TIMS.GetListValue(listTechID2)
                    'If v_listTechID2 <> "" AndAlso TIMS.VAL1(txtPHour.Text) > 0 Then
                    '    If CHK_TechID_PHOUR_Exceed_54_Hour(pms_CHK_SAVE_TrainDescRserve, String.Concat("T2x", v_listTechID2), TIMS.VAL1(txtPHour.Text)) Then
                    '        Dim txt_listTechID2 As String=TIMS.GetListText(listTechID2)
                    '        Dim s_ERR1 As String=String.Concat(txt_listTechID2, "-", "授課時數不得超過54小時。(", v_listTechID2, ")", vbCrLf)
                    '        If errMsg.IndexOf(s_ERR1)=-1 Then errMsg &= s_ERR1
                    '    End If
                    'End If

                Case Cst_i科場地
                    listPTID.Style("background-color") = "#FFFFFF"
                    If listPTID.SelectedIndex = 0 Then
                        errMsg &= "尚有上課地點未選擇。" & vbCrLf
                        listPTID.Style("background-color") = "#FFCCFF"
                    End If

                Case Cst_i上課時間
                    txtPName.Text = TIMS.ClearSQM(txtPName.Text)
                    txtPName.Text = txtPName.Text.Replace(" ", "").Replace("　", "")
                    txtPName.Style("background-color") = "#FFFFFF"
                    If txtPName.Text = "" Then
                        errMsg &= "尚有授課時間未輸入(請勿只填入空白)。" & vbCrLf
                        txtPName.Style("background-color") = "#FFCCFF"
                    End If
                    '18:30~21:30
                    Dim flag_chk_pname As Boolean = TIMS.Check_PName(txtPName.Text)
                    If Not flag_chk_pname Then
                        errMsg &= String.Format("授課時間格式有誤(填入範例：09:00~12:00)。 {0}", txtPName.Text) & vbCrLf
                        txtPName.Style("background-color") = "#FFCCFF"
                    End If

                    txtPHour.Text = TIMS.ClearSQM(txtPHour.Text)
                    txtPHour.Style("background-color") = "#FFFFFF"
                    If txtPHour.Text = "" OrElse Not IsNumeric(txtPHour.Text) Then
                        errMsg &= "尚有時數輸入錯誤，限輸入數字。" & vbCrLf
                        txtPHour.Style("background-color") = "#FFCCFF"
                    ElseIf IsNumeric(txtPHour.Text) AndAlso TIMS.CINT1(txtPHour.Text) <= 0 Then
                        errMsg &= "尚有時數輸入錯誤，必須大於0。" & vbCrLf
                        txtPHour.Style("background-color") = "#FFCCFF"
                    End If

                    txtEHour.Text = TIMS.ClearSQM(txtEHour.Text)
                    txtEHour.Style("background-color") = "#FFFFFF"
                    If txtEHour.Text <> "" AndAlso Not IsNumeric(txtEHour.Text) Then
                        errMsg &= "技檢訓練時數輸入錯誤，限輸入數字。" & vbCrLf
                        txtEHour.Style("background-color") = "#FFCCFF"
                    ElseIf txtEHour.Text <> "" AndAlso IsNumeric(txtEHour.Text) AndAlso TIMS.CINT1(txtEHour.Text) <= 0 Then
                        errMsg &= "技檢訓練時數輸入錯誤，必須大於0。" & vbCrLf
                        txtEHour.Style("background-color") = "#FFCCFF"
                    ElseIf txtEHour.Text <> "" AndAlso IsNumeric(txtEHour.Text) AndAlso IsNumeric(txtPHour.Text) AndAlso TIMS.CINT1(txtEHour.Text) > TIMS.CINT1(txtPHour.Text) Then
                        errMsg &= "技檢訓練時數輸入錯誤，「該欄位數字只能等於或小於」訓練時數。" & vbCrLf
                        txtEHour.Style("background-color") = "#FFCCFF"
                    End If

                    '授課時段'早上'下午'晚上
                    Dim TPERIOD28_Style_background_color As String = "#FFFFFF"
                    Dim sTPERIOD28 As String = $"{If(TPERIOD28_1t.Checked, "Y", "N")}{If(TPERIOD28_2t.Checked, "Y", "N")}{If(TPERIOD28_3t.Checked, "Y", "N")}"
                    If sTPERIOD28 = cst_NNN Then
                        errMsg &= "授課時段:早上、下午、晚上 至少要設定其中一項" & vbCrLf
                        TPERIOD28_Style_background_color = "#FFCCFF"
                    ElseIf TIMS.CHK_STR_CNT(sTPERIOD28, "Y") > 1 Then
                        errMsg &= "授課時段:早上、下午、晚上為單選"
                        TPERIOD28_Style_background_color = "#FFCCFF"
                    End If
                    TPERIOD28_1t.Style("background-color") = TPERIOD28_Style_background_color
                    TPERIOD28_2t.Style("background-color") = TPERIOD28_Style_background_color
                    TPERIOD28_3t.Style("background-color") = TPERIOD28_Style_background_color

                Case Cst_i課程表
                    Dim s_ABSDate As String = TIMS.Cdate3(If(ASDate.Text <> "", ASDate.Text, If(BSDate.Text <> "", BSDate.Text, STDate.Value)))
                    Dim s_ABEDate As String = TIMS.Cdate3(If(AEDate.Text <> "", AEDate.Text, If(BEDate.Text <> "", BEDate.Text, FDDate.Value)))
                    If txtSTrainDate.Text <> "" AndAlso IsDate(txtSTrainDate.Text) Then
                        txtSTrainDate.Style("background-color") = "#FFFFFF"
                        If CDate(txtSTrainDate.Text) < CDate(s_ABSDate) OrElse CDate(txtSTrainDate.Text) > CDate(s_ABEDate) Then
                            errMsg &= "日期未完全輸入 或 輸入區間錯誤(超出訓練期間範圍)。" & vbCrLf
                            txtSTrainDate.Style("background-color") = "#FFCCFF"
                        End If
                    Else
                        errMsg &= "日期未完全輸入 或 輸入區間錯誤(超出訓練期間範圍)。" & vbCrLf
                        txtSTrainDate.Style("background-color") = "#FFCCFF"
                    End If
                    listTechID.Style("background-color") = "#FFFFFF"
                    If listTechID.SelectedIndex = 0 Then
                        errMsg &= "尚有任課教師未選擇。" & vbCrLf
                        listTechID.Style("background-color") = "#FFCCFF"
                    End If
                    listTechID2.Style("background-color") = "#FFFFFF"
                    If listTechID2.SelectedIndex = 0 Then
                        'errMsg &= "尚有任課教師未選擇。" & vbCrLf
                        listTechID2.Style("background-color") = "#FFCCFF"
                    End If
                    listPTID.Style("background-color") = "#FFFFFF"
                    If listPTID.SelectedIndex = 0 Then
                        errMsg &= "尚有上課地點未選擇。" & vbCrLf
                        listPTID.Style("background-color") = "#FFCCFF"
                    End If
                    txtPName.Text = TIMS.ClearSQM(txtPName.Text)
                    txtPName.Text = txtPName.Text.Replace(" ", "").Replace("　", "")
                    txtPName.Style("background-color") = "#FFFFFF"
                    If txtPName.Text = "" Then
                        errMsg &= "尚有授課時間未輸入(請勿只填入空白)。" & vbCrLf
                        txtPName.Style("background-color") = "#FFCCFF"
                    End If
                    '18:30~21:30
                    Dim flag_chk_pname As Boolean = TIMS.Check_PName(txtPName.Text)
                    If Not flag_chk_pname Then
                        errMsg &= String.Format("授課時間格式有誤(填入範例：09:00~12:00)。 {0}", txtPName.Text) & vbCrLf
                        txtPName.Style("background-color") = "#FFCCFF"
                    End If

                    txtPHour.Text = TIMS.ClearSQM(txtPHour.Text)
                    txtPHour.Style("background-color") = "#FFFFFF"
                    If txtPHour.Text = "" OrElse Not IsNumeric(txtPHour.Text) Then
                        errMsg &= "尚有時數輸入錯誤，限輸入數字。" & vbCrLf
                        txtPHour.Style("background-color") = "#FFCCFF"
                    ElseIf IsNumeric(txtPHour.Text) AndAlso TIMS.CINT1(txtPHour.Text) <= 0 Then
                        errMsg &= "尚有時數輸入錯誤，必須大於0。" & vbCrLf
                        txtPHour.Style("background-color") = "#FFCCFF"
                    End If
                    txtEHour.Text = TIMS.ClearSQM(txtEHour.Text)
                    txtEHour.Style("background-color") = "#FFFFFF"
                    If txtEHour.Text <> "" AndAlso Not IsNumeric(txtEHour.Text) Then
                        errMsg &= "技檢訓練時數輸入錯誤，限輸入數字。" & vbCrLf
                        txtEHour.Style("background-color") = "#FFCCFF"
                    ElseIf txtEHour.Text <> "" AndAlso IsNumeric(txtEHour.Text) AndAlso TIMS.CINT1(txtEHour.Text) <= 0 Then
                        errMsg &= "技檢訓練時數輸入錯誤，必須大於0。" & vbCrLf
                        txtEHour.Style("background-color") = "#FFCCFF"
                    ElseIf txtEHour.Text <> "" AndAlso IsNumeric(txtEHour.Text) AndAlso IsNumeric(txtPHour.Text) AndAlso TIMS.CINT1(txtEHour.Text) > TIMS.CINT1(txtPHour.Text) Then
                        errMsg &= "技檢訓練時數輸入錯誤，「該欄位數字只能等於或小於」訓練時數。" & vbCrLf
                        txtEHour.Style("background-color") = "#FFCCFF"
                    End If

                    '授課時段'早上'下午'晚上
                    Dim TPERIOD28_Style_background_color As String = "#FFFFFF"
                    Dim sTPERIOD28 As String = $"{If(TPERIOD28_1t.Checked, "Y", "N")}{If(TPERIOD28_2t.Checked, "Y", "N")}{If(TPERIOD28_3t.Checked, "Y", "N")}"
                    If sTPERIOD28 = cst_NNN Then
                        errMsg &= "授課時段:早上、下午、晚上 至少要設定其中一項" & vbCrLf
                        TPERIOD28_Style_background_color = "#FFCCFF"
                    ElseIf TIMS.CHK_STR_CNT(sTPERIOD28, "Y") > 1 Then
                        errMsg &= "授課時段:早上、下午、晚上為單選"
                        TPERIOD28_Style_background_color = "#FFCCFF"
                    End If
                    TPERIOD28_1t.Style("background-color") = TPERIOD28_Style_background_color
                    TPERIOD28_2t.Style("background-color") = TPERIOD28_Style_background_color
                    TPERIOD28_3t.Style("background-color") = TPERIOD28_Style_background_color

                Case Cst_i遠距教學
                    '※系統檢核：當【變更內容】選擇"申請整班為遠距教學"時，請檢核所有課程，都有勾選【遠距教學】。
                    '班級-遠距教學
                    'vrbl_DISTANCE  null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
                    If vrbl_DISTANCE.Equals("1") AndAlso Not bx_FARLEARN.Checked Then
                        errMsg &= String.Format("【變更內容】選擇「{0}」時，所有課程(表)，都須勾選【遠距教學】欄位", s_DISTANCE_N) & vbCrLf
                        Exit For
                    ElseIf vrbl_DISTANCE.Equals("3") AndAlso bx_FARLEARN.Checked Then
                        errMsg &= String.Format("【變更內容】選擇「{0}」時，所有課程(表)，都不須勾選【遠距教學】欄位", s_DISTANCE_N) & vbCrLf
                        Exit For
                    End If
                    i_DG3ROWS += 1
                    If (bx_FARLEARN.Checked) Then i_FARLEARN += 1
            End Select
        Next
        '班級-遠距教學
        'vrbl_DISTANCE  null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        'Dim s_log1 As String=""
        's_log1 &= String.Format(" v_ChgItem: {0}", v_ChgItem) & vbCrLf
        's_log1 &= String.Format(" vrbl_DISTANCE: {0}", vrbl_DISTANCE) & vbCrLf
        's_log1 &= String.Format(" i_DG3ROWS: {0}", i_DG3ROWS) & vbCrLf
        's_log1 &= String.Format(" i_FARLEARN: {0}", i_FARLEARN) & vbCrLf
        's_log1 &= String.Format("v_ChgItem.Equals(CStr(Cst_i遠距教學)): {0}", v_ChgItem.Equals(CStr(Cst_i遠距教學))) & vbCrLf
        'TIMS.LOG.Debug(s_log1)

        If v_ChgItem.Equals(CStr(Cst_i遠距教學)) AndAlso vrbl_DISTANCE.Equals("2") Then
            'Dim s_DISTANCE_N As String=TIMS.GET_DISTANCE_N(0, vrbl_DISTANCE)
            If i_DG3ROWS.Equals(i_FARLEARN) Then
                errMsg &= String.Format("【變更內容】選擇「{0}」時，不能將所有課程(表)都勾選【遠距教學】欄位", s_DISTANCE_N) & vbCrLf
            ElseIf i_FARLEARN = 0 Then
                errMsg &= String.Format("【變更內容】選擇「{0}」時，請於課程(表)勾選【遠距教學】欄位", s_DISTANCE_N) & vbCrLf
            End If
        End If
        'Common.MessageBox(Me, errMsg)
        If errMsg <> "" Then Return errMsg '異常

        Dim objConn As SqlConnection = DbAccess.GetConnection()
        Call TIMS.OpenDbConn(objConn)
        Dim objTrans As SqlTransaction = objConn.BeginTransaction() '使用預設隔離等級 'IsolationLevel.ReadCommitted

        Try
            'Dim objConn As SqlConnection
            'TIMS.TestDbConn(Me, objConn, True)
            'objConn.Open()
            'IsolationLevel.ReadCommitted
            'objTrans=objConn.BeginTransaction("ReviseItem")
            'objTrans=objConn.BeginTransaction() '使用預設隔離等級 'IsolationLevel.ReadCommitted

            '有舊的PTDRID，所以直接進行Update。
            Dim flag_oldItem As Boolean = True
            If iPTDRID = 0 Then  '編輯狀態就沒取得PTDRID時，代表是新增，所以取得新的PTDRID。
                flag_oldItem = False '新增
                'INSERT PLAN_TRAINDESC_REVISE
                iPTDRID = INSERT_PLANTRAINDESCREVISE(rPlanID, rComIDNO, rSeqNo, Today.ToString("yyyy/MM/dd"), ViewState(vs_SubSeqNO), objConn, objTrans)
            End If

            'PLAN_TRAINDESC
            For Each eItem As DataGridItem In DataGrid3.Items
                Dim hidePTDRID As HtmlInputHidden = eItem.FindControl("hide_PTDRID")
                Dim hidePTDID As HtmlInputHidden = eItem.FindControl("hide_PTDID")
                Dim hideID1 As HtmlInputHidden = eItem.FindControl("hide_ID1")
                Dim hideID2 As HtmlInputHidden = eItem.FindControl("hide_ID2")
                Dim hideID3 As HtmlInputHidden = eItem.FindControl("hide_ID3") '時數
                Dim hideID9 As HtmlInputHidden = eItem.FindControl("hide_ID9") '技檢訓練時數

                Dim hideID4 As HtmlInputHidden = eItem.FindControl("hide_ID4")
                Dim hideID5 As HtmlInputHidden = eItem.FindControl("hide_ID5")
                Dim txtSTrainDate As TextBox = eItem.FindControl("txt_STrainDate")
                Dim txtPName As TextBox = eItem.FindControl("txt_PName")
                Dim txtPHour As TextBox = eItem.FindControl("txt_PHour") '時數
                Dim txtEHour As TextBox = eItem.FindControl("txt_EHour") '技檢訓練時數
                Dim listPTID As DropDownList = eItem.FindControl("list_PTID")
                Dim listTechID As DropDownList = eItem.FindControl("list_TechID")
                Dim hideID6 As HtmlInputHidden = eItem.FindControl("hide_ID6")
                Dim listTechID2 As DropDownList = eItem.FindControl("list_TechID2")
                Dim hideSTrainDate As HtmlInputHidden = eItem.FindControl("hide_STrainDate")
                Dim hidePName As HtmlInputHidden = eItem.FindControl("hide_PName")
                Dim hidePHour As HtmlInputHidden = eItem.FindControl("hide_PHour") '時數
                Dim hideEHour As HtmlInputHidden = eItem.FindControl("hide_EHour") '技檢訓練時數
                Dim hidePTID As HtmlInputHidden = eItem.FindControl("hide_PTID")
                Dim hideTechID As HtmlInputHidden = eItem.FindControl("hide_TechID")
                Dim hideTechID2 As HtmlInputHidden = eItem.FindControl("hide_TechID2")

                Dim TPERIOD28_1t As CheckBox = eItem.FindControl("TPERIOD28_1t")
                Dim TPERIOD28_2t As CheckBox = eItem.FindControl("TPERIOD28_2t")
                Dim TPERIOD28_3t As CheckBox = eItem.FindControl("TPERIOD28_3t")
                Dim hidTPERIOD28 As HtmlInputHidden = eItem.FindControl("hidTPERIOD28")
                Dim hideID7 As HtmlInputHidden = eItem.FindControl("hide_ID7")

                'Cst_i遠距教學
                Dim bx_FARLEARN As CheckBox = eItem.FindControl("bx_FARLEARN")
                Dim hide_FARLEARN As HtmlInputHidden = eItem.FindControl("hide_FARLEARN")
                Dim hideID8 As HtmlInputHidden = eItem.FindControl("hide_ID8")
                Dim sFARLEARN As String = If(bx_FARLEARN.Checked, "Y", "")
                '授課時段'早上'下午'晚上
                Dim sTPERIOD28 As String = String.Concat(If(TPERIOD28_1t.Checked, "Y", "N"), If(TPERIOD28_2t.Checked, "Y", "N"), If(TPERIOD28_3t.Checked, "Y", "N"))

                hideSTrainDate.Value = TIMS.Cdate3(hideSTrainDate.Value)
                txtSTrainDate.Text = TIMS.Cdate3(txtSTrainDate.Text)
                If flag_oldItem Then
                    '有申請過變更
                    Select Case v_ChgItem'.SelectedValue '產投用選項
                        Case Cst_i訓練期間
                            If hideID1.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID1.Value, 1, txtSTrainDate.Text, objConn, objTrans)
                        Case Cst_i師資
                            If hideID5.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID5.Value, 5, listTechID.SelectedValue, objConn, objTrans)
                        Case Cst_i助教
                            If hideID6.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID6.Value, 6, listTechID2.SelectedValue, objConn, objTrans)
                        Case Cst_i科場地
                            If hideID4.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID4.Value, 4, listPTID.SelectedValue, objConn, objTrans)
                        Case Cst_i上課時間
                            If hideID2.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID2.Value, 2, txtPName.Text, objConn, objTrans)
                            If hideID3.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID3.Value, 3, txtPHour.Text, objConn, objTrans)
                            If hideID9.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID9.Value, 9, txtEHour.Text, objConn, objTrans)
                            If hideID7.Value <> "" Then Call Update_PlanTrainDescReviseItem(hideID7.Value, 7, sTPERIOD28, objConn, objTrans)
                        Case Cst_i遠距教學
                            If hideID8.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 8, hide_FARLEARN.Value, sFARLEARN, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID8.Value, 8, sFARLEARN, objConn, objTrans)
                            End If
                        Case Cst_i課程表
                            If hideID1.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 1, hideSTrainDate.Value, txtSTrainDate.Text, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID1.Value, 1, txtSTrainDate.Text, objConn, objTrans)
                            End If
                            If hideID2.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 2, hidePName.Value, txtPName.Text, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID2.Value, 2, txtPName.Text, objConn, objTrans)
                            End If
                            '時數
                            If hideID3.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 3, hidePHour.Value, txtPHour.Text, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID3.Value, 3, txtPHour.Text, objConn, objTrans)
                            End If
                            '技檢訓練時數
                            If hideID9.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 9, hideEHour.Value, txtEHour.Text, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID9.Value, 9, txtEHour.Text, objConn, objTrans)
                            End If
                            If hideID7.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 7, hidTPERIOD28.Value, sTPERIOD28, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID7.Value, 7, sTPERIOD28, objConn, objTrans)
                            End If
                            If hideID8.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 8, hide_FARLEARN.Value, sFARLEARN, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID8.Value, 8, sFARLEARN, objConn, objTrans)
                            End If
                            If hideID4.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 4, hidePTID.Value, listPTID.SelectedValue, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID4.Value, 4, listPTID.SelectedValue, objConn, objTrans)
                            End If
                            If hideID5.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 5, hideTechID.Value, listTechID.SelectedValue, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID5.Value, 5, listTechID.SelectedValue, objConn, objTrans)
                            End If
                            If hideID6.Value = "" Then
                                Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 6, hideTechID2.Value, listTechID2.SelectedValue, objConn, objTrans)
                            Else
                                Call Update_PlanTrainDescReviseItem(hideID6.Value, 6, listTechID2.SelectedValue, objConn, objTrans)
                            End If
                    End Select
                Else
                    '沒申請過變更
                    Select Case v_ChgItem'.SelectedValue '產投用選項
                        Case Cst_i訓練期間
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 1, hideSTrainDate.Value, txtSTrainDate.Text, objConn, objTrans)
                        Case Cst_i師資
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 5, hideTechID.Value, listTechID.SelectedValue, objConn, objTrans)
                        Case Cst_i助教
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 6, hideTechID2.Value, listTechID2.SelectedValue, objConn, objTrans)
                        Case Cst_i科場地
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 4, hidePTID.Value, listPTID.SelectedValue, objConn, objTrans)
                        Case Cst_i上課時間
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 2, hidePName.Value, txtPName.Text, objConn, objTrans)
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 3, hidePHour.Value, txtPHour.Text, objConn, objTrans) 'PHour'時數
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 9, hideEHour.Value, txtEHour.Text, objConn, objTrans) 'EHour'技檢訓練時數
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 7, hidTPERIOD28.Value, sTPERIOD28, objConn, objTrans)
                        Case Cst_i遠距教學
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 8, hide_FARLEARN.Value, sFARLEARN, objConn, objTrans)
                        Case Cst_i課程表
                            'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 1, hideSTrainDate.Value, txtSTrainDate.Text, objConn, objTrans) 'STrainDate
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 2, hidePName.Value, txtPName.Text, objConn, objTrans) 'PName
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 3, hidePHour.Value, txtPHour.Text, objConn, objTrans) 'PHour'時數
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 9, hideEHour.Value, txtEHour.Text, objConn, objTrans) 'EHour'技檢訓練時數
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 7, hidTPERIOD28.Value, sTPERIOD28, objConn, objTrans)
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 8, hide_FARLEARN.Value, sFARLEARN, objConn, objTrans)
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 4, hidePTID.Value, listPTID.SelectedValue, objConn, objTrans) 'PTID
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 5, hideTechID.Value, listTechID.SelectedValue, objConn, objTrans) 'TechID
                            Call Insert_PlanTrainDescReviseItem(hidePTDID.Value, iPTDRID, ApplyItem, 6, hideTechID2.Value, listTechID2.SelectedValue, objConn, objTrans) 'TechID2
                    End Select
                End If
            Next
            objTrans.Commit()
            rst = ""
            'objConn.Close()
            'objTrans.Dispose()
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Common.MessageBox(Me, String.Concat(ex.Message, "-75B6"))
            objTrans.Rollback() 'Rollback("ReviseItem")
            Call TIMS.CloseDbConn(objConn)
            If flagDebugTest Then Throw ex
        End Try
        Call TIMS.CloseDbConn(objConn)
        Return rst
    End Function

    ''' <summary>'超過54小數為TRUE 沒超過 FALSE</summary>
    ''' <param name="pms_CHK1"></param>
    ''' <param name="vTeachID"></param>
    ''' <param name="iPHour"></param>
    ''' <returns></returns>
    Private Function CHK_TechID_PHOUR_Exceed_54_Hour(ByRef pms_CHK1 As Hashtable, vTeachID As String, iPHour As Double) As Boolean
        Dim ALL_HOUR As String = TIMS.GetMyValue2(pms_CHK1, vTeachID) '取得目前值，可能為空
        Dim iALL_HOUR As Integer = TIMS.VAL1(ALL_HOUR) '處理數值:若為空=0
        If iPHour > 0 Then iALL_HOUR += iPHour '(目前要加入的時數)
        TIMS.SetMyValue2(pms_CHK1, vTeachID, iALL_HOUR) '儲存總時數
        Return (iALL_HOUR > 54)
    End Function

    ''' <summary>正式送出 產學訓 進入儲存階段2 (課程表儲存) 儲存 目前只有產投有2階段儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub But_Save28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Save28.Click
        'Dim blnOkSave1 As Boolean=Save_TrainDescRserve(ViewState(vs_UpdateItemIndex))
        Dim strOkSave1 As String = CHK_SAVE_TrainDescRserve(ViewState(vs_UpdateItemIndex))
        If strOkSave1 <> "" Then
            Common.MessageBox(Page, String.Concat("計畫變更申請(課程表)失敗!!", vbCrLf, strOkSave1))
            Exit Sub
        End If

        ViewState(vs_UpdateTrainDesc) = "Y"
        Try
            Call Save_Sub_TPlanID28()
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Dim exMessage As String = ex.Message
            Common.MessageBox(Page, "計畫變更申請(課程表)失敗!!" & exMessage)
            'Throw ex 'Common.MessageBox(Page, ex.ToString)
        End Try


    End Sub

    '(確定) andy  產學訓 進入儲存階段1
    Private Sub But_Sub28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Sub28.Click
        '非產投計畫(TIMS)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            'TIMS
            But_Sub.Style("display") = "none" '一般計畫儲存
            But_Sub28.Style("display") = "none" '產投計畫確定
            But_Save28.Style("display") = "none" '產投計畫儲存
            'But_UPLOAD28.Style("display")="none" '產投計畫儲存(上傳檔案)
            Common.MessageBox(Me, String.Concat(TIMS.cst_NODATAMsg3, "-33F5"))
            Exit Sub
        End If

        txtIntaxno.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txtIntaxno.Text))
        txtUname.Text = TIMS.ClearSQM(txtUname.Text)
        Dim v_ChgItem As String = TIMS.GetListValue(ChgItem)
        Dim v_PackageTypeNew As String = TIMS.GetListValue(PackageTypeNew)
        Select Case v_ChgItem 'ChgItem.SelectedValue
            Case Cst_i包班種類
                Dim ErrMsg As String = ""
                Select Case v_PackageTypeNew'PackageTypeNew.SelectedValue
                    Case "2" '充電起飛計畫 '企業包班
                        If txtUname.Text = "" Then
                            txtUname.Text = ""
                            ErrMsg &= "包班事業單位 企業名稱，不可為空" & vbCrLf
                        Else
                            '錯誤檢查 txtUname.Text = Trim(txtUname.Text)
                            If txtUname.Text.Length > 50 Then ErrMsg &= "包班事業單位 企業名稱，長度超過限制範圍50文字長度" & vbCrLf
                        End If
                        If txtIntaxno.Text <> "" Then
                            'txtIntaxno.Text = Trim(txtIntaxno.Text)
                            If Not TIMS.CheckIsECFA(txtIntaxno.Text, gobjconn) Then ErrMsg &= $"「{txtUname.Text}」該包班事業單位 企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf '未填寫 ECFA包班事業單位資料
                        Else
                            txtIntaxno.Text = ""
                            ErrMsg &= "包班事業單位 服務單位統一編號，不可為空" & vbCrLf
                        End If
                        If ErrMsg = "" Then If Not Session("Revise_BusPackage") Is Nothing Then Session("Revise_BusPackage") = Nothing '沒有錯誤清空session 
                    Case "3" '充電起飛計畫 '聯合企業包班
                        'If Session("Revise_BusPackage") Is Nothing Then
                        'End If
                        If Not Session("Revise_BusPackage") Is Nothing Then
                            Dim j As Integer = 0
                            Dim dt1 As DataTable = Session("Revise_BusPackage")
                            If dt1.Rows.Count > 0 Then
                                For i As Integer = 0 To dt1.Rows.Count - 1
                                    If Not dt1.Rows(i).RowState = DataRowState.Deleted Then
                                        Dim drX As DataRow = dt1.Rows(i)
                                        If Not TIMS.CheckIsECFA($"{drX("Intaxno")}", gobjconn) Then ErrMsg &= $"「{drX("Uname")}」該包班事業單位 企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf '未填寫 包班事業單位資料
                                        j += 1
                                    End If
                                Next
                                If j = 0 Then ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf '未填寫 包班事業單位資料
                            Else
                                ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf '未填寫 包班事業單位資料
                            End If
                        Else
                            ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf '未填寫 包班事業單位資料
                        End If
                    Case Else
                        'If Not Session("Revise_BusPackage") Is Nothing Then Session("Revise_BusPackage")=Nothing
                End Select
                If ErrMsg <> "" Then
                    Common.SetListItem(PackageTypeNew, "3")
                    Common.MessageBox(Me, ErrMsg)
                    Exit Sub
                End If
                Call Save_Sub_TPlanID28()

            Case Cst_i師資 '產投用選項
                ViewState(vs_UpdateItemIndex) = Cst_i師資
                Call checkSubSeqNo()
                If TeacherName1_2.Text = "" Then
                    Common.MessageBox(Page, "變更師資尚未選擇")
                    Exit Sub
                End If
                ViewState(vs_UpdateTrainDesc) = "N"
                Call Save_Sub_TPlanID28()
                Button6.Style("display") = "none"
                Call HindItem()
                Call CreateTrainDesc()

            Case Cst_i助教 '產投用選項
                ViewState(vs_UpdateItemIndex) = Cst_i助教
                Call checkSubSeqNo()
                If TeacherName2_2.Text = "" Then
                    Common.MessageBox(Page, "變更助教尚未選擇")
                    Exit Sub
                End If
                ViewState(vs_UpdateTrainDesc) = "N"
                Call Save_Sub_TPlanID28()
                Button6_2.Style("display") = "none"
                Call HindItem()
                Call CreateTrainDesc()

            Case Cst_i科場地
                Dim v_NewData14_1b As String = TIMS.GetListValue(NewData14_1b) '學科場地地址1
                Dim v_NewData14_2b As String = TIMS.GetListValue(NewData14_2b) '術科場地地址1
                Dim v_NewData14_3 As String = TIMS.GetListValue(NewData14_3) '學科場地地址2
                Dim v_NewData14_4 As String = TIMS.GetListValue(NewData14_4) '術科場地地址2
                Dim drSciPc1 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_1b, 1, gobjconn) '取得學科場地的地址
                Dim drTechPc1 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_2b, 2, gobjconn) '取得術科場地的地址
                Dim drSciPc2 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_3, 1, gobjconn) '取得學科場地的地址2
                Dim drTechPc2 As DataRow = TIMS.Get_SciTechDR(rComIDNO, v_NewData14_4, 2, gobjconn) '取得術科場地的地址2
                Hid_NewData8_4.Value = If((drSciPc1 IsNot Nothing), drSciPc1("PTID").ToString(), "") '學科場地地址
                Hid_NewData8_5.Value = If((drTechPc1 IsNot Nothing), drTechPc1("PTID").ToString(), "") '術科場地地址
                Hid_NewData8_6.Value = If((drSciPc2 IsNot Nothing), drSciPc2("PTID").ToString(), "") '學科場地地址2
                Hid_NewData8_7.Value = If((drTechPc2 IsNot Nothing), drTechPc2("PTID").ToString(), "") '術科場地地址2
                Dim strErrmsg As String = ""
                If v_NewData14_3 <> "" AndAlso v_NewData14_1b = "" Then strErrmsg &= "學科場地2 有選值，學科場地1(不可為空)!" & vbCrLf
                If v_NewData14_4 <> "" AndAlso v_NewData14_2b = "" Then strErrmsg &= "術科場地2 有選值，術科場地1(不可為空)!" & vbCrLf
                If v_NewData14_1b = "" AndAlso v_NewData14_2b = "" Then strErrmsg &= "學科場地1 或 術科場地1 (不可為空)!" & vbCrLf
                If v_NewData14_3 <> "" AndAlso v_NewData14_1b <> "" AndAlso v_NewData14_3 = v_NewData14_1b Then strErrmsg &= "學科場地1 與 學科場地2(不可為相同)!" & vbCrLf
                If v_NewData14_4 <> "" AndAlso v_NewData14_2b <> "" AndAlso v_NewData14_4 = v_NewData14_2b Then strErrmsg &= "術科場地1 與 術科場地2(不可為相同)!" & vbCrLf
                If strErrmsg <> "" Then
                    Common.MessageBox(Me, strErrmsg)
                    Exit Sub
                End If

                '當學術科場地沒有選取時，先檢查課程表中是否有學術科項目，有的話，就不能沒選。
                Dim flag_checkCF1 As Boolean = If(v_NewData14_1b = "", Check_PlanTrainDescClassification(rPlanID, rComIDNO, rSeqNo, "1", gobjconn), False)
                Dim flag_checkCF2 As Boolean = If(v_NewData14_2b = "", Check_PlanTrainDescClassification(rPlanID, rComIDNO, rSeqNo, "2", gobjconn), False)
                Dim flag_checkCF3 As Boolean = If(v_NewData14_3 = "", Check_PlanTrainDescClassification(rPlanID, rComIDNO, rSeqNo, "1", gobjconn), False)
                Dim flag_checkCF4 As Boolean = If(v_NewData14_4 = "", Check_PlanTrainDescClassification(rPlanID, rComIDNO, rSeqNo, "2", gobjconn), False)
                Dim tmpMsg As String = String.Empty
                If flag_checkCF1 AndAlso flag_checkCF3 Then tmpMsg = "因課程表中有學科課程，所以不能沒有選取學科場地。" & vbCrLf
                If flag_checkCF2 AndAlso flag_checkCF4 Then tmpMsg &= "因課程表中有術科課程，所以不能沒有選取術科場地。"
                'If TaddressS2.SelectedIndex <= 0 And TaddressT2.SelectedIndex <= 0 Then tmpMsg += "【學科上課地址】、【術科上課地址】，必須至少設定一項。" & vbCrLf
                If tmpMsg <> "" Then
                    Common.MessageBox(Me, tmpMsg)
                    Exit Sub
                End If

                ViewState(vs_UpdateItemIndex) = Cst_i科場地
                Call checkSubSeqNo()
                ViewState(vs_UpdateTrainDesc) = "N"
                Button6.Style("display") = "none"
                'NewData14_1.Enabled=False
                'NewData14_2.Enabled=False
                NewData14_1b.Enabled = False
                NewData14_2b.Enabled = False
                NewData14_3.Enabled = False
                NewData14_4.Enabled = False
                'TaddressS2.Enabled=False
                'TaddressT2.Enabled=False
                Call HindItem()
                Call CreateTrainDesc()

            Case Cst_i上課時間
                If DataGrid2.Items.Count = 0 Then
                    Common.MessageBox(Page, "變更內容-請輸入欲變更之上課時間!-CE66")
                    Exit Sub
                End If
                ViewState(vs_UpdateItemIndex) = Cst_i上課時間
                Call checkSubSeqNo()
                ViewState(vs_UpdateTrainDesc) = "N"
                Call HindItem()
                Call CreateTrainDesc()

            Case Cst_i訓練期間
                ViewState(vs_UpdateItemIndex) = Cst_i訓練期間
                Call checkSubSeqNo()
                If ASDate.Text = "" OrElse AEDate.Text = "" Then
                    Common.MessageBox(Me, "請輸入變更內容的起迄日期!")
                    Exit Sub
                End If
                If CDate(ASDate.Text) = CDate(BSDate.Text) AndAlso CDate(AEDate.Text) = CDate(BEDate.Text) Then
                    Common.MessageBox(Me, "新的訓練日期不能與舊日期的相同!")
                    Exit Sub
                End If
                Call Save_Sub_TPlanID28()
                If ViewState(vs_UpdateTrainDesc) = "N" Then
                    IMG1.Style("display") = "none"
                    IMG2.Style("display") = "none"
                    ASDate.Enabled = False
                    AEDate.Enabled = False
                    Call HindItem()
                    Call CreateTrainDesc()
                End If

            Case Cst_i遠距教學
                Dim vrbl_DISTANCE As String = TIMS.GetListValue(rbl_DISTANCE)
                If vrbl_DISTANCE = "" Then
                    Common.MessageBox(Me, "遠距教學 請選擇 變更內容!")
                    Exit Sub
                End If
                ViewState(vs_UpdateItemIndex) = Cst_i遠距教學
                Call checkSubSeqNo()
                ViewState(vs_UpdateTrainDesc) = "N"
                Call HindItem()
                Call CreateTrainDesc()

            Case Cst_i課程表
                ViewState(vs_UpdateItemIndex) = Cst_i課程表
                Call checkSubSeqNo()
                ViewState(vs_UpdateTrainDesc) = "N"
                Call HindItem()
                Call CreateTrainDesc()
            Case 0
                Exit Select
            Case Else
                Datagrid3Table.Style("display") = "none"
                Call Save_Sub_TPlanID28()
        End Select
    End Sub

    ''' <summary> 回傳 ViewState(vs_SubSeqNO) 值 </summary>
    Sub checkSubSeqNo()
        Dim R_SubSeqNO As Object = Nothing
        ViewState(vs_SubSeqNO) = 1 '預設為1

        Try
            Dim i_AltDataID As Integer = If(ViewState(vs_UpdateItemIndex) IsNot Nothing, Val(ViewState(vs_UpdateItemIndex)), -1)
            Dim objstr As String = ""
            objstr = " SELECT MAX(SubSeqNo) MaxSubSeqNo FROM PLAN_REVISE WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
            objstr &= " AND CDate=" & TIMS.To_date(ApplyDate.Text)
            If i_AltDataID > 0 Then objstr &= " AND AltDataID=" & ViewState(vs_UpdateItemIndex)
            R_SubSeqNO = DbAccess.ExecuteScalar(objstr, gobjconn)
            If Not IsDBNull(R_SubSeqNO) Then
                '取得目前值
                ViewState(vs_SubSeqNO) = CInt(R_SubSeqNO)
            Else
                '項目為空 取得最大值+1
                objstr = " SELECT MAX(SubSeqNo) MaxSubSeqNo FROM PLAN_REVISE WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'"
                objstr &= " and CDate=" & TIMS.To_date(ApplyDate.Text)
                R_SubSeqNO = DbAccess.ExecuteScalar(objstr, gobjconn)
                If Not IsDBNull(R_SubSeqNO) Then ViewState(vs_SubSeqNO) = CInt(R_SubSeqNO) + 1
            End If
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)
            'If flagDebugTest Then Throw ex '有錯就停止吧
            Throw ex '有錯就停止吧
        End Try
    End Sub

    ''' <summary>隱藏 But_Sub28 SHOW But_Save28</summary>
    Sub HindItem()
        'Dim TD_1 As HtmlTableCell=FindControl("TD_1")
        'Dim TD_2 As HtmlTableCell=FindControl("TD_2")
        ChgItem.Enabled = False
        TIMS.Tooltip(ChgItem, "狀態鎖定")
        ApplyDate.Enabled = False
        imgApplyDate.Visible = False
        TIMS.Tooltip(ApplyDate, "狀態鎖定")
        But_Sub28.Style("display") = "none"
        But_Save28.Style("display") = cst_inline1
        'But_UPLOAD28.Style("display")=cst_inline1 '"none" '產投計畫儲存(上傳檔案)
        TD_1.Style("display") = cst_inline1
        TD_2.Style("display") = cst_inline1
    End Sub

    '20081023 andy 變更上課時段，所選日期未排課時新增一筆資料(日期須在訓練期間內)
    Sub AddClass_Sch(ByVal OCID As String, ByVal NewSchDay As String, ByVal msg As Label)
        '產投無課程資料 / TIMS才有排課資訊
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Exit Sub
        '非產投計畫(TIMS)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Dim sErrmsg As String = ""
            Dim sqlstr As String = ""
            Dim dt1 As DataTable = Nothing
            Dim dr1 As DataRow = Nothing

            Try
                If Not CheckClassSchedule(ViewState(vs_OCID), sErrmsg, gobjconn) Then
                    If hid_chkmsg.Value = "on" Then
                        msg.Text = sErrmsg
                        hid_chkmsg.Value = "off"
                        Common.MessageBox(Page, sErrmsg)
                    End If
                    Exit Sub
                Else
                    sqlstr = " SELECT MAX(CONVERT(NUMERIC, OCID)) OCID, MAX(Type) Type ,MAX(Formal) Formal FROM CLASS_SCHEDULE WHERE OCID=@OCID "
                    Dim sCmd As New SqlCommand(sqlstr, gobjconn)
                    Call TIMS.OpenDbConn(gobjconn)
                    dt1 = New DataTable
                    With sCmd
                        .Parameters.Clear()
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
                        dt1.Load(.ExecuteReader())
                    End With
                    If dt1.Rows.Count = 0 Then
                        sErrmsg = "本班目前尚未排課，請先確認是否已於課程管理完成排課作業！-2370"
                        msg.Text = sErrmsg
                        hid_chkmsg.Value = "off"
                        Common.MessageBox(Page, sErrmsg)
                        Exit Sub
                    End If
                    dr1 = dt1.Rows(0)
                End If

                sqlstr = " SELECT * FROM CLASS_SCHEDULE WHERE OCID=@OCID AND SchoolDate=@SchoolDate"
                Dim sCmd2 As New SqlCommand(sqlstr, gobjconn)
                Call TIMS.OpenDbConn(gobjconn)
                dt1 = New DataTable
                With sCmd2
                    .Parameters.Clear()
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
                    .Parameters.Add("SchoolDate", SqlDbType.DateTime).Value = Convert.ToDateTime(NewSchDay)
                    dt1.Load(.ExecuteReader())
                End With

#Region "(No Use)"

                'If dt1.Rows.Count=0 Then
                '    Dim sql As String=""
                '    sql="" & vbCrLf
                '    'sql &= " /* IDENTITY(1,1):  CSID */" & vbCrLf
                '    sql &= " INSERT INTO CLASS_SCHEDULE(" & vbCrLf
                '    sql &= " CSID" & vbCrLf
                '    sql &= " ,OCID" & vbCrLf
                '    sql &= " ,SchoolDate" & vbCrLf
                '    sql &= " ,Type" & vbCrLf
                '    sql &= " ,ModifyAcct" & vbCrLf
                '    sql &= " ,ModifyDate" & vbCrLf
                '    sql &= " ,Formal" & vbCrLf
                '    sql &= " ) VALUES (" & vbCrLf
                '    sql &= " @CSID" & vbCrLf
                '    sql &= " ,@OCID" & vbCrLf
                '    sql &= " ,@SchoolDate" & vbCrLf
                '    sql &= " ,@Type" & vbCrLf
                '    sql &= " ,@ModifyAcct" & vbCrLf
                '    sql &= " ,getdate()" & vbCrLf
                '    sql &= " ,@Formal" & vbCrLf
                '    sql &= " )" & vbCrLf
                '    Dim iCmd As New SqlCommand(sql, gobjconn)
                '    Call TIMS.OpenDbConn(gobjconn)
                '    Dim iCSID As Integer=DbAccess.GetNewId(gobjconn, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                '    With iCmd
                '        .Parameters.Clear()
                '        .Parameters.Add("CSID", SqlDbType.Int).Value=iCSID
                '        .Parameters.Add("OCID", SqlDbType.VarChar).Value=OCID
                '        ' .Parameters.Add("SchoolDate", SqlDbType.VarChar).Value=CDate(NewSchDay).toString("yyyy/MM/dd")
                '        .Parameters.Add("SchoolDate", SqlDbType.DateTime).Value=Convert.ToDateTime(NewSchDay)
                '        .Parameters.Add("Type", SqlDbType.VarChar).Value=Convert.ToString(dr1("Type"))
                '        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value=Convert.ToString(sm.UserInfo.UserID)
                '        .Parameters.Add("Formal", SqlDbType.VarChar).Value=Convert.ToString(dr1("Formal"))
                '        .ExecuteNonQuery()
                '    End With
                'End If

#End Region
            Catch ex As Exception
                Dim strErrmsg5 As String = ""
                strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
                strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
                Call TIMS.WriteTraceLog(strErrmsg5)
                Common.MessageBox(Me, String.Concat(ex.Message, "-6E04"))
                If flagDebugTest Then Throw ex
            End Try
        End If
    End Sub

    '20080902 andy  add     
    Sub CreateClassList(ByVal TimeClass As String, ByVal SDate As String, ByVal EDate As String)
        'Dim ClassArray As New ArrayList
        Dim dt As New DataTable
        dt.Clear()
        '定義DataTable欄位將資料帶入listbox 中
        dt.Columns.Add("ClassNO", Type.GetType("System.Decimal"))
        dt.Columns.Add("ClassListName", Type.GetType("System.String"))
        Dim sqlstr As String = ""
        If TimeClass = "start" Then ViewState(vs_IsLoaded) = "Y" '是否已選擇過開始日期 

        '判斷是否該班級已有排課
        'sqlstr=" select  ocid   from CLASS_SCHEDULE where  OCID=" & ViewState(vs_OCID)
        'dt3=DbAccess.GetDataTable(sqlstr)
        'If dt3.Rows.Count=0 Then
        '    If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        If hid_chkmsg.Value="on" Then
        '            hid_chkmsg.Value="off"
        '            Common.MessageBox(Page, "本班目前尚未排課，請先確認是否已於課程管理完成排課作業！")
        '        End If
        '    End If
        '    Exit Function
        'End If

        '產投無課程資料 / TIMS才有排課資訊
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Exit Sub

        Dim sErrmsg As String = ""
        If Not CheckClassSchedule(ViewState(vs_OCID), sErrmsg, gobjconn) Then
            'msg.Text=sErrmsg
            If hid_chkmsg.Value = "on" Then
                hid_chkmsg.Value = "off"
                Common.MessageBox(Page, sErrmsg)
            End If
            Exit Sub
        End If
        '先取出兩個日期共有的堂數是安排在那幾堂課,將得到結果加入ClassArray  
        'sqlstr="" & vbCrLf
        'sqlstr += "  select max(Class1) as Class1 , max(Class2) as Class2 , max(Class3) as Class3 , max(Class4) as Class4" & vbCrLf
        'sqlstr += " ,max(Class5) as Class5 , max(Class6) as Class6 , max(Class7) as Class7 , max(Class8) as Class8" & vbCrLf
        'sqlstr += " ,max(Class9) as Class9 ,max(Class10) as Class10,max(Class11) as Class11,max(Class12) as Class12" & vbCrLf
        sqlstr = ""
        For i As Integer = 1 To 12
            sqlstr &= String.Format(If((i = 1), " SELECT class{0}", " ,class{0}"), CStr(i))
        Next
        sqlstr &= " FROM CLASS_SCHEDULE" & vbCrLf
        sqlstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
        sqlstr &= " AND SchoolDate=" & TIMS.To_date(SDate) & vbCrLf
        Dim drS As DataRow = DbAccess.GetOneRow(sqlstr, gobjconn)

        sqlstr = ""
        For i As Integer = 1 To 12
            sqlstr &= String.Format(If((i = 1), " SELECT class{0}", " ,class{0}"), CStr(i))
        Next
        sqlstr &= " FROM CLASS_SCHEDULE" & vbCrLf
        sqlstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
        sqlstr &= " AND SchoolDate=" & TIMS.To_date(EDate) & vbCrLf
        Dim drE As DataRow = DbAccess.GetOneRow(sqlstr, gobjconn)
        If drS Is Nothing AndAlso drE Is Nothing Then
            Common.MessageBox(Page, "查無資料(該班的開結訓日期範圍" & TRange.Text & ")")
            'msg5.Text="欲更換日期兩日皆無安排課程！"
            'hid_chklist.Value="N"
            Exit Sub
        End If

        Dim i_S1 As Integer = 0 '非空堂
        Dim i_E1 As Integer = 0 '非空堂
        If Not drS Is Nothing Then
            For i As Integer = 1 To 12
                If Not drS Is Nothing Then
                    If Convert.ToString(drS("Class" & i)) <> "" Then
                        If Val(drS("Class" & i)) > 0 Then i_S1 += 1 '非空堂
                    End If
                End If
                If Not drE Is Nothing Then
                    If Convert.ToString(drE("Class" & i)) <> "" Then
                        If Val(drE("Class" & i)) > 0 Then i_E1 += 1 '非空堂
                    End If
                End If
            Next
        End If

        hid_chklist.Value = "Y"
        msg5.Text = ""
        If i_S1 = 0 AndAlso i_E1 = 0 Then
            msg5.Text = "(原計畫內容／變更內容) 欲更換日期兩日皆無安排課程！"
            hid_chklist.Value = "N"
        ElseIf i_S1 = 0 Then
            msg5.Text = "(原計畫內容) 原日期未安排課程！"
        ElseIf i_E1 = 0 Then
            msg5.Text = "(變更內容) 變更日未安排課程！"
        End If

        'ClassArray.Clear()
        'Dim j As Integer=0
        'For i As Integer=1 To 12
        '    '090413 and  edit  訓練時段 改為就算是空堂也show出來
        '    If IsDBNull(dr("Class" & i)) Then j += 1
        '    ClassArray.Add(i)
        'Next
        'hid_chklist.Value="Y"
        'If j=12 And SDate <> "" And EDate <> "" Then
        '    msg5.Text="欲更換日期兩日皆無安排課程！"
        '    hid_chklist.Value="N"
        'ElseIf j=12 And (SDate <> "" And EDate="") Then
        '    msg5.Text="原日期未安排課程！"
        'ElseIf j=12 And (SDate="" And EDate <> "") Then
        '    msg5.Text="變更日未安排課程！"
        'Else
        '    msg5.Text=""
        'End If

        If ViewState(vs_OCID) <> "" Then
            sqlstr = ""
            sqlstr &= " SELECT * FROM CLASS_SCHEDULE" & vbCrLf
            sqlstr &= " WHERE OCID=" & ViewState(vs_OCID) & " "
            Select Case TimeClass
                Case "start"
                    sqlstr &= " and SchoolDate=" & TIMS.To_date(SDate)
                Case "end"
                    sqlstr &= " and SchoolDate=" & TIMS.To_date(EDate)
            End Select
            Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, gobjconn)

            If Not dr Is Nothing Then
                'Dim j As Integer=0
                For i As Integer = 1 To 12
                    If IsDBNull(dr("Class" & i)) Then
                        'Dim m As Integer
                        'For Each m As Integer In ClassArray
                        '    If i=m Then
                        '        Dim dr2 As DataRow=dt.NewRow()
                        '        dr2("ClassNO")=i
                        '        dr2("ClassListName")=("第" & i & "節--(未排課)")
                        '        dt.Rows.Add(dr2)
                        '    End If
                        'Next
                        Dim dr2 As DataRow = dt.NewRow()
                        dr2("ClassNO") = i
                        dr2("ClassListName") = ("第" & i & "節--(未排課)")
                        dt.Rows.Add(dr2)
                    Else
                        Dim dr2 As DataRow = dt.NewRow()
                        dr2("ClassNO") = i 'Dim sClassListName As String="" 'sClassListName=""
                        Dim sClassListName As String = String.Concat("第", i, "節--", If(IsDBNull(dr("Class" & i)), "", TIMS.Get_CourseName(dr("Class" & i), Nothing, gobjconn)))
                        sClassListName &= String.Concat("--", If(IsDBNull(dr("Teacher" & i)), "", TIMS.Get_TeachCName(dr("Teacher" & i), gobjconn)))
                        sClassListName &= CStr(If(IsDBNull(dr("Teacher" & i + 12)), "", (" " & TIMS.Get_TeachCName(dr("Teacher" & i + 12), gobjconn))))
                        'TIMS.Get_TeachCName(drv("TechID2"), objconn) '
                        dr2("ClassListName") = sClassListName
                        dt.Rows.Add(dr2)
                    End If
                Next
            End If
        End If
        'If TimeClass="start" Then
        '    dsClasslist.Tables.Add(dt)
        'ElseIf TimeClass="end" Then
        '    dsClasslist2.Tables.Add(dt)
        'End If

        If TimeClass <> "" Then
            Select Case TimeClass
                Case "start"
                    SourceLB1.DataTextField = "ClassListName"
                    SourceLB1.DataValueField = "ClassNO"
                    SourceLB1.DataSource = dt 'dsClasslist
                    SourceLB1.DataBind()
                Case "end"
                    SourceLB2.DataTextField = "ClassListName"
                    SourceLB2.DataValueField = "ClassNO"
                    SourceLB2.DataSource = dt 'dsClasslist2
                    SourceLB2.DataBind()
            End Select
        End If
    End Sub

    ''檢查該班是否已結訓
    'Function CheckIsClosed(ByVal OCID As String) As Boolean
    '    Dim sql As String
    '    Dim result1 As String
    '    'IsClosed	是否結訓
    '    sql=""
    '    sql &= " SELECT IsClosed FROM CLASS_CLASSINFO"
    '    sql &= " WHERE IsSuccess='Y' AND NotOpen='N'"
    '    sql &= " AND PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'"
    '    sql &= " AND ocid='" & OCID & "') "
    '    result1=DbAccess.ExecuteScalar(sql, gobjconn)

    '    If result1.Trim()="Y" Then
    '        Return True '已結訓
    '    Else
    '        Return False
    '    End If
    'End Function

    '課程名稱
    Function ShowClassList(ByVal orderlist As String, ByVal courlist As String, ByVal kind As String) As String
        Dim itemstr As String = ""
        Dim listary As String() = Split(orderlist, ",")
        Dim courary As String() = Split(courlist, ",")
        If kind = "name" Then
            itemstr = ""
            For i As Integer = 0 To listary.Length - 1
                If (courary.Length <> listary.Length) And courary.Length = 1 Then '由於舊版程式帶出來的OldData2_2  課程只有一個值無法對應新程式
                    itemstr = itemstr & "(" & CStr(listary(i)) & ") " & CStr(TIMS.Get_CourseName(courary(0), Nothing, gobjconn)) & "  "
                Else
                    If (CStr(courary(i)) = "x") Or (Trim(CStr(courary(i))) = "") Then
                        itemstr &= "(" & CStr(listary(i)) & ") " & "未排課" & "  "
                    Else
                        itemstr &= "(" & CStr(listary(i)) & ") " & CStr(TIMS.Get_CourseName(courary(i), Nothing, gobjconn)) & "  "
                    End If
                End If
            Next
            itemstr = Left(itemstr, itemstr.Length - 2)       '課程名稱
        ElseIf kind = "item" Then
            itemstr = ""
            For i As Integer = 0 To listary.Length - 1
                If itemstr <> "" Then itemstr &= ","
                itemstr &= itemstr & CStr(listary(i))
            Next
        End If
        Return itemstr
    End Function

    Private Sub btnAdd_1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnAdd_1.Click
        If SourceLB2.Items.Count + TargetLB2.Items.Count = 0 Then
            Common.MessageBox(Page, " 變更日期尚未選取無法進行設定!")
            Exit Sub
        End If
        Dim i As Integer = 0
        While i <= SourceLB1.Items.Count - 1
            If SourceLB1.Items(i).Selected Then
                TargetLB1.Items.Add(New ListItem(SourceLB1.Items(i).Text, SourceLB1.Items(i).Value))
                SourceLB1.Items.Remove(SourceLB1.Items(i))
            Else
                i += 1
            End If
        End While
    End Sub

    Private Sub btnAddAll_1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnAddAll_1.Click
        If SourceLB2.Items.Count + TargetLB2.Items.Count = 0 Then
            Common.MessageBox(Page, "變更日期尚未選取無法進行設定!")
            Exit Sub
        End If
        If (SourceLB1.Items.Count > 0) Then
            Dim classLi As New ListItem
            For Each classLi In SourceLB1.Items
                TargetLB1.Items.Add(classLi)
            Next
        End If
        SourceLB1.Items.Clear()
    End Sub

    Private Sub btnRemove_1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnRemove_1.Click
        If SourceLB2.Items.Count + TargetLB2.Items.Count = 0 Then
            Common.MessageBox(Page, "變更日期尚未選取無法進行設定!")
            Exit Sub
        End If
        Dim i As Integer = 0
        While i <= TargetLB1.Items.Count - 1
            If TargetLB1.Items(i).Selected Then
                SourceLB1.Items.Add(New ListItem(TargetLB1.Items(i).Text, TargetLB1.Items(i).Value))
                TargetLB1.Items.Remove(TargetLB1.Items(i))
            Else
                i += 1
            End If
        End While
    End Sub

    Private Sub btnRemoveAll_1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnRemoveAll_1.Click
        If SourceLB2.Items.Count + TargetLB2.Items.Count = 0 Then
            Common.MessageBox(Page, "變更日期尚未選取無法進行設定!")
            Exit Sub
        End If
        If (TargetLB1.Items.Count > 0) Then
            Dim classLi As New ListItem
            For Each classLi In TargetLB1.Items
                SourceLB1.Items.Add(classLi)
            Next
        End If
        TargetLB1.Items.Clear()
    End Sub

    Private Sub btnAdd_2_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnAdd_2.Click
        Dim i As Integer = 0
        While i <= SourceLB2.Items.Count - 1
            If SourceLB2.Items(i).Selected Then
                TargetLB2.Items.Add(New ListItem(SourceLB2.Items(i).Text, SourceLB2.Items(i).Value))
                SourceLB2.Items.Remove(SourceLB2.Items(i))
            Else
                i += 1
            End If
        End While
    End Sub

    Private Sub btnAddAll_2_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnAddAll_2.Click
        If (SourceLB2.Items.Count > 0) Then
            Dim classLi As New ListItem
            For Each classLi In SourceLB2.Items
                TargetLB2.Items.Add(classLi)
            Next
        End If
        SourceLB2.Items.Clear()
    End Sub

    Private Sub btnRemove_2_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnRemove_2.Click
        Dim i As Integer = 0
        While i <= TargetLB2.Items.Count - 1
            If TargetLB2.Items(i).Selected Then
                SourceLB2.Items.Add(New ListItem(TargetLB2.Items(i).Text, TargetLB2.Items(i).Value))
                TargetLB2.Items.Remove(TargetLB2.Items(i))
            Else
                i += 1
            End If
        End While
    End Sub

    Private Sub btnRemoveAll_2_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles btnRemoveAll_2.Click
        If (TargetLB2.Items.Count > 0) Then
            'Dim classLi As New ListItem
            For Each classLi As ListItem In TargetLB2.Items
                SourceLB2.Items.Add(classLi)
            Next
        End If
        TargetLB2.Items.Clear()
    End Sub

    Private Function Get_ClassRow(ByVal sOCID As String, ByVal sSchoolDate As String, ByRef tConn As SqlConnection, ByRef tTrans As SqlTransaction) As DataRow
        Dim Rst As DataRow = Nothing
        'Dim da As SqlDataAdapter=TIMS.GetOneDA()
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM CLASS_SCHEDULE "
        sql &= " WHERE OCID='" & sOCID & "'"
        sql &= " AND SchoolDate=" & TIMS.To_date(sSchoolDate)
        Dim dt As New DataTable
        'Call TIMS.OpenDbConn(tConn)
        Dim sCmd As New SqlCommand(sql, tConn, tTrans)
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then Rst = dt.Rows(0)
        'Rst=TIMS.GetOneRow(sql, da)
        Return Rst
    End Function

#Region "課程表Function"

    ''' <summary>'目前課表('變更申請前)</summary>
    ''' <param name="tmpPlanID"></param>
    ''' <param name="tmpComIDNO"></param>
    ''' <param name="tmpSeqNO"></param>
    ''' <returns></returns>
    Private Function Get_PlanTrainDesc(ByVal tmpPlanID As Integer, ByVal tmpComIDNO As String, ByVal tmpSeqNO As Integer) As DataTable
        Dim Rst As New DataTable '= Nothing
        'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        'PLAN_TRAINDESC '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        Dim s_Parms As New Hashtable From {{"planid", tmpPlanID}, {"comidno", tmpComIDNO}, {"seqno", tmpSeqNO}}
        Dim sql As String = ""
        sql &= " SELECT NULL AS PTDRID ,pd.PTDID" & vbCrLf
        sql &= " ,NULL AS ID1" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, pd.STrainDate, 111) STrainDate" & vbCrLf
        sql &= " ,NULL AS ID2 ,pd.PName" & vbCrLf
        sql &= " ,NULL AS ID3 ,pd.PHour" & vbCrLf '時數
        sql &= " ,NULL AS ID9 ,pd.EHour" & vbCrLf '技檢訓練時數
        sql &= " ,NULL AS ID7 ,pd.TPERIOD28" & vbCrLf
        sql &= " ,NULL AS ID8 ,pd.FARLEARN" & vbCrLf 'ID8
        sql &= " ,pd.PCont" & vbCrLf
        sql &= " ,pd.Classification1" & vbCrLf
        sql &= " ,NULL AS ID4 ,pd.PTID" & vbCrLf
        sql &= " ,NULL AS ID5" & vbCrLf
        sql &= " ,NULL AS ID6" & vbCrLf
        sql &= " ,pd.TechID" & vbCrLf
        sql &= " ,pd.TechID2" & vbCrLf
        sql &= " ,pp.TechPlaceID TechPlaceID" & vbCrLf
        sql &= " ,pp.TechPlaceID2 TechPlaceID2" & vbCrLf
        sql &= " ,pp.SciPlaceID SciPlaceID" & vbCrLf
        sql &= " ,pp.SciPlaceID2 SciPlaceID2" & vbCrLf

        sql &= " FROM PLAN_TRAINDESC pd" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.PlanID=pd.PlanID AND pp.ComIDNO=pd.ComIDNO AND pp.SeqNO=pd.SeqNO" & vbCrLf
        sql &= " WHERE pd.PlanID=@planid AND pd.ComIDNO=@comidno AND pd.SeqNO=@seqno" & vbCrLf
        'sql &= " AND pd.AltDataID=@AltDataID" & vbCrLf
        sql &= " ORDER BY pd.STrainDate ,pd.PName ASC" & vbCrLf

        Rst = DbAccess.GetDataTable(sql, gobjconn, s_Parms)
        Return Rst
    End Function

    ''' <summary>'課程表申請變更前</summary>
    ''' <param name="tmpID"></param>
    ''' <param name="sAltDataID"></param>
    ''' <param name="sType"></param>
    ''' <returns></returns>
    Private Function Get_PlanTrainDescOldRevise(ByVal tmpID As Integer, ByVal sAltDataID As String, ByVal sType As String) As DataTable
        'sType: 'sType@now新版課表(產投) 'sType@old1舊1課表(產投)
        Dim rst As New DataTable '= Nothing
        'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        'Dim da As SqlDataAdapter=TIMS.GetOneDA(gobjconn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.PTDRID ,g.PTDID" & vbCrLf
        sql &= " ,b.PTDRIID AS ID1 ,CONVERT(VARCHAR, ISNULL(CONVERT(DATETIME, b.OldData, 111) ,g.STrainDate), 111) STrainDate" & vbCrLf
        sql &= " ,c.PTDRIID AS ID2 ,ISNULL(c.OldData,g.PName) AS PName" & vbCrLf
        sql &= " ,d.PTDRIID AS ID3 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d.OldData),g.PHour)) AS PHour" & vbCrLf '時數
        sql &= " ,d9.PTDRIID AS ID9 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d9.OldData),g.EHour)) AS EHour" & vbCrLf '技檢訓練時數
        sql &= " ,d7.PTDRIID AS ID7 ,ISNULL(d7.OldData,g.TPERIOD28) TPERIOD28" & vbCrLf
        sql &= " ,d8.PTDRIID AS ID8 ,ISNULL(d8.OldData,g.FARLEARN) FARLEARN" & vbCrLf
        sql &= " ,g.PCont" & vbCrLf
        sql &= " ,g.Classification1" & vbCrLf
        sql &= " ,e.PTDRIID AS ID4 ,ISNULL(e.OldData,g.PTID) AS PTID" & vbCrLf
        sql &= " ,f.PTDRIID AS ID5 ,ISNULL(f.OldData,g.TechID) AS TechID" & vbCrLf
        sql &= " ,f2.PTDRIID AS ID6 ,ISNULL(f2.OldData,g.TechID2) AS TechID2" & vbCrLf
        sql &= " FROM PLAN_TRAINDESC_REVISE a" & vbCrLf
        Select Case sType
            Case cst_now
                sql &= " JOIN PLAN_TRAINDESC_RO g ON g.PlanID=a.PlanID AND g.ComIDNO=a.ComIDNO AND g.SeqNO=a.SeqNO AND g.PTDRID=a.PTDRID AND a.PTDRID=@PTDRid" & vbCrLf
            Case Else 'cst_old1
                sql &= " JOIN PLAN_TRAINDESC g ON g.PlanID=a.PlanID AND g.ComIDNO=a.ComIDNO AND g.SeqNO=a.SeqNO AND a.PTDRID=@PTDRid" & vbCrLf
        End Select
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM b ON b.PTDID=g.PTDID AND b.PTDRID=a.PTDRID AND b.AltDataItem=1 AND b.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM c ON c.PTDID=g.PTDID AND c.PTDRID=a.PTDRID AND c.AltDataItem=2 AND c.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d ON d.PTDID=g.PTDID AND d.PTDRID=a.PTDRID AND d.AltDataItem=3 AND d.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM e ON e.PTDID=g.PTDID AND e.PTDRID=a.PTDRID AND e.AltDataItem=4 AND e.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f ON f.PTDID=g.PTDID AND f.PTDRID=a.PTDRID AND f.AltDataItem=5 AND f.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f2 ON f2.PTDID=g.PTDID AND f2.PTDRID=a.PTDRID AND f2.AltDataItem=6 and f2.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d7 ON d7.PTDID=g.PTDID AND d7.PTDRID=a.PTDRID AND d7.AltDataItem=7 AND d7.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d8 ON d8.PTDID=g.PTDID AND d8.PTDRID=a.PTDRID AND d8.AltDataItem=8 AND d8.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d9 ON d9.PTDID=g.PTDID AND d9.PTDRID=a.PTDRID AND d9.AltDataItem=9 AND d9.AltDataID=@AltDataID" & vbCrLf
        'sql &= " order by g.STrainDate,g.PName" & vbCrLf
        'sql &= " order by STrainDate,PName" & vbCrLf
        Dim sCmd As New SqlCommand(sql, gobjconn)
        Call TIMS.OpenDbConn(gobjconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRid", SqlDbType.Int).Value = tmpID
            .Parameters.Add("AltDataID", SqlDbType.VarChar).Value = sAltDataID
            rst.Load(.ExecuteReader())
        End With
        rst.DefaultView.Sort = "STrainDate,PName"
        rst = TIMS.dv2dt(rst.DefaultView)
        Return rst
    End Function

    ''' <summary>'課程表申請變更後</summary>
    ''' <param name="iPTDRID"></param>
    ''' <param name="sAltDataID"></param>
    ''' <returns></returns>
    Private Function Get_PlanTrainDescNewRevise(ByVal iPTDRID As Integer, ByVal sAltDataID As String) As DataTable
        Dim rst As New DataTable '= Nothing
        'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        Dim sqlStr As String = ""
        sqlStr = ""
        sqlStr &= " SELECT a.PTDRID ,g.PTDID" & vbCrLf
        sqlStr &= " ,b.PTDRIID AS ID1 ,CONVERT(VARCHAR, ISNULL(CONVERT(DATETIME, LTRIM(rtrim(b.NewData)), 111) ,g.STrainDate), 111) STrainDate" & vbCrLf
        sqlStr &= " ,c.PTDRIID AS ID2 ,ISNULL(c.NewData,g.PName) PName" & vbCrLf
        'sqlStr &= " ,d.PTDRIID AS ID3 ,ISNULL(CONVERT(NUMERIC, LTRIM(RTRIM(d.NewData))),g.PHour) as PHour" & vbCrLf
        sqlStr &= " ,d.PTDRIID AS ID3 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d.NewData),g.PHour)) AS PHour" & vbCrLf '時數
        sqlStr &= " ,d9.PTDRIID AS ID9 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d9.NewData),g.EHour)) AS EHour" & vbCrLf '技檢訓練時數
        sqlStr &= " ,d7.PTDRIID AS ID7 ,ISNULL(d7.NewData,g.TPERIOD28) TPERIOD28" & vbCrLf
        sqlStr &= " ,d8.PTDRIID AS ID8 ,ISNULL(d8.NewData,g.FARLEARN) FARLEARN" & vbCrLf
        sqlStr &= " ,g.PCont" & vbCrLf
        sqlStr &= " ,g.Classification1" & vbCrLf
        sqlStr &= " ,e.PTDRIID AS ID4 ,ISNULL((e.NewData),g.PTID) AS PTID" & vbCrLf
        sqlStr &= " ,f.PTDRIID AS ID5 ,ISNULL(LTRIM(RTRIM(f.NewData)),g.TechID) AS TechID" & vbCrLf
        sqlStr &= " ,f2.PTDRIID AS ID6 ,ISNULL(LTRIM(RTRIM(f2.NewData)),g.TechID2) AS TechID2" & vbCrLf
        sqlStr &= " ,pp.SciPlaceID ,pp.SciPlaceID2 ,pp.TechPlaceID ,pp.TechPlaceID2" & vbCrLf
        sqlStr &= " FROM PLAN_PLANINFO pp" & vbCrLf
        sqlStr &= " JOIN PLAN_TRAINDESC_REVISE a ON pp.PlanID=a.PlanID AND pp.ComIDNO=a.ComIDNO AND pp.SeqNO=a.SeqNO AND a.PTDRID=@PTDRID" & vbCrLf
        sqlStr &= " JOIN PLAN_TRAINDESC g ON g.PlanID=a.PlanID AND g.ComIDNO=a.ComIDNO AND g.SeqNO=a.SeqNO" & vbCrLf
        'PTDRIID
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM b ON b.PTDID=g.PTDID AND b.PTDRID=a.PTDRID AND b.AltDataItem=1 AND b.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM c ON c.PTDID=g.PTDID AND c.PTDRID=a.PTDRID AND c.AltDataItem=2 AND c.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d ON d.PTDID=g.PTDID AND d.PTDRID=a.PTDRID AND d.AltDataItem=3 AND d.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM e ON e.PTDID=g.PTDID AND e.PTDRID=a.PTDRID AND e.AltDataItem=4 AND e.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f ON f.PTDID=g.PTDID AND f.PTDRID=a.PTDRID AND f.AltDataItem=5 AND f.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f2 ON f2.PTDID=g.PTDID AND f2.PTDRID=a.PTDRID AND f2.AltDataItem=6 AND f2.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d7 ON d7.PTDID=g.PTDID AND d7.PTDRID=a.PTDRID AND d7.AltDataItem=7 AND d7.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d8 ON d8.PTDID=g.PTDID AND d8.PTDRID=a.PTDRID AND d8.AltDataItem=8 AND d8.AltDataID=@AltDataID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d9 ON d9.PTDID=g.PTDID AND d9.PTDRID=a.PTDRID AND d9.AltDataItem=9 AND d9.AltDataID=@AltDataID" & vbCrLf

        'sqlStr += " ORDER BY g.STrainDate ,PName "
        'sqlStr += " ORDER BY STrainDate ,PName "
        Dim sCmd As New SqlCommand(sqlStr, gobjconn)
        Call TIMS.OpenDbConn(gobjconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRID", SqlDbType.BigInt).Value = iPTDRID
            .Parameters.Add("AltDataID", SqlDbType.VarChar).Value = sAltDataID
            rst.Load(.ExecuteReader())
        End With
        rst.DefaultView.Sort = "STrainDate,PName"
        rst = TIMS.dv2dt(rst.DefaultView)
        Return rst
    End Function

    ''' <summary> 取得可能序號，當天一種 只能申請一次，刪除異常資料。('iType :1 為申請 /2 為變更結果) </summary>
    ''' <param name="iType"></param>
    ''' <param name="tmpID"></param>
    ''' <param name="tmpCOMIDNO"></param>
    ''' <param name="tmpSNO"></param>
    ''' <param name="tmpDate"></param>
    ''' <param name="tmpSubSNO"></param>
    ''' <param name="sAltDataID"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Function Get_PTDRID(ByVal iType As Integer, ByVal tmpID As Integer, ByVal tmpCOMIDNO As String, ByVal tmpSNO As Integer,
                        ByVal tmpDate As String, ByVal tmpSubSNO As Integer, ByVal sAltDataID As String, ByRef oConn As SqlConnection) As Integer
        'iType :1 為申請 /2 為變更結果
        'sAltDataID  : 申請@PLAN_PLANINFO:為空  變更結果@PLAN_REVISE:有值
        Dim iRst As Integer = 0
        Dim sql As String = ""
        'Dim oConn As SqlConnection=DbAccess.GetConnection() 'Call TIMS.OpenDbConn(oConn)
        '整理異常資料並刪除
        Dim flag_PTDRID_OK As Boolean = Get_PTDRIDxDelErr(sm, iType, tmpID, tmpCOMIDNO, tmpSNO, tmpDate, tmpSubSNO, sAltDataID, oConn)

        Select Case iType
            Case 1
                If Not flag_PTDRID_OK Then
                    Common.MessageBox(Me, cst_errmsg_alt5)
                    Return iRst
                End If

                '應該正常資料
                sql = "" & vbCrLf
                sql &= " SELECT p.PTDRID" & vbCrLf
                sql &= " FROM PLAN_TRAINDESC_REVISE p" & vbCrLf
                sql &= " WHERE p.PlanID=@planid AND p.ComIDNO=@comidno AND p.SeqNO=@seqno" & vbCrLf
                sql &= " AND p.CDate=@cdate AND p.SubSeqNO=@subseqno" & vbCrLf
                'sql &= "   AND p.ptdrid IN (" & tmpPTDRIDs & ")" & vbCrLf
                Dim sCmd2 As New SqlCommand(sql, oConn)
                Dim dt2 As New DataTable
                With sCmd2
                    .Parameters.Clear()
                    .Parameters.Add("planid", SqlDbType.Int).Value = TIMS.GetValue1(tmpID)
                    .Parameters.Add("comidno", SqlDbType.VarChar).Value = TIMS.GetValue1(tmpCOMIDNO)
                    .Parameters.Add("seqno", SqlDbType.Int).Value = TIMS.GetValue1(tmpSNO)
                    .Parameters.Add("cdate", SqlDbType.DateTime).Value = CDate(tmpDate)
                    .Parameters.Add("subseqno", SqlDbType.Int).Value = TIMS.GetValue1(tmpSubSNO)
                    '.Parameters.Add("AltDataID", SqlDbType.VarChar).Value=sAltDataID
                    'Call TIMS.Upd_NULL(sCmd2)
                    dt2.Load(.ExecuteReader())
                End With
                If dt2.Rows.Count > 0 Then
                    If Not IsDBNull(dt2.Rows(0)("PTDRID")) Then iRst = Val(dt2.Rows(0)("PTDRID"))
                End If

                sql = "" & vbCrLf
                sql &= " SELECT x.PTDRID" & vbCrLf
                sql &= " FROM PLAN_TRAINDESC_REVISEITEM x" & vbCrLf
                sql &= " WHERE x.AltDataID=@AltDataID AND x.PTDRID=@PTDRID "
                Dim sCmd1 As New SqlCommand(sql, oConn)
                Dim dt1 As New DataTable
                With sCmd1
                    .Parameters.Clear()
                    .Parameters.Add("AltDataID", SqlDbType.VarChar).Value = sAltDataID
                    .Parameters.Add("PTDRID", SqlDbType.VarChar).Value = iRst
                    dt1.Load(.ExecuteReader())
                End With
                If dt1.Rows.Count = 0 Then iRst = 0 '不應該為無資料(清除資訊)

            Case 2
                '應該正常資料
                sql = "" & vbCrLf
                sql &= " SELECT MAX(p.PTDRID) PTDRID" & vbCrLf
                sql &= " FROM PLAN_TRAINDESC_REVISE p" & vbCrLf
                sql &= " WHERE p.PlanID=@planid AND p.ComIDNO=@comidno AND p.SeqNO=@seqno" & vbCrLf
                sql &= " AND p.CDate=@cdate AND p.SubSeqNO=@subseqno" & vbCrLf
                'sql &= "   AND p.ptdrid IN (" & tmpPTDRIDs & ")" & vbCrLf
                Dim sCmd2 As New SqlCommand(sql, oConn)
                Dim dt2 As New DataTable
                With sCmd2
                    .Parameters.Clear()
                    .Parameters.Add("planid", SqlDbType.Int).Value = TIMS.GetValue1(tmpID)
                    .Parameters.Add("comidno", SqlDbType.VarChar).Value = TIMS.GetValue1(tmpCOMIDNO)
                    .Parameters.Add("seqno", SqlDbType.Int).Value = TIMS.GetValue1(tmpSNO)
                    .Parameters.Add("cdate", SqlDbType.DateTime).Value = CDate(tmpDate)
                    .Parameters.Add("subseqno", SqlDbType.Int).Value = TIMS.GetValue1(tmpSubSNO)
                    '.Parameters.Add("AltDataID", SqlDbType.VarChar).Value=sAltDataID
                    'Call TIMS.Upd_NULL(sCmd2)
                    dt2.Load(.ExecuteReader())
                End With
                If dt2.Rows.Count > 0 Then
                    If Not IsDBNull(dt2.Rows(0)("PTDRID")) Then iRst = Val(dt2.Rows(0)("PTDRID"))
                End If

                sql = "" & vbCrLf
                sql &= " SELECT x.PTDRID" & vbCrLf
                sql &= " FROM PLAN_TRAINDESC_REVISEITEM x" & vbCrLf
                sql &= " WHERE x.AltDataID=@AltDataID" & vbCrLf
                sql &= " AND x.PTDRID=@PTDRID AND x.PTDRIID IS NOT NULL" 'PTDRIID
                Dim sCmd1 As New SqlCommand(sql, oConn)
                Dim dt1 As New DataTable
                With sCmd1
                    .Parameters.Clear()
                    .Parameters.Add("AltDataID", SqlDbType.VarChar).Value = sAltDataID
                    .Parameters.Add("PTDRID", SqlDbType.VarChar).Value = iRst
                    dt1.Load(.ExecuteReader())
                End With
                If dt1.Rows.Count = 0 Then iRst = 0 '不應該為無資料(清除資訊)
        End Select
        'Call TIMS.CloseDbConn(oConn)
        Return iRst
    End Function

    ''' <summary> 檢核是否有課程申請資料 </summary>
    ''' <param name="ptdrid"></param>
    ''' <param name="iAltDataID"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Function Check_PlanTrainDescReviseItem(ByVal ptdrid As Integer, ByVal iAltDataID As Integer, ByVal oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT 'x' FROM PLAN_TRAINDESC_REVISEITEM "
        sql &= " WHERE PTDRID=@PTDRID "
        sql &= " AND AltDataID=@AltDataID AND PTDID IS NOT NULL " 'PTDID
        Dim sCmd As New SqlCommand(sql, oConn)
        TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRID", SqlDbType.Int).Value = ptdrid
            .Parameters.Add("AltDataID", SqlDbType.Int).Value = iAltDataID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    ''' <summary> UPDATE PLAN_TRAINDESC_REVISEITEM </summary>
    ''' <param name="ptdriid">流水號</param>
    ''' <param name="altdataitem">儲存詳資序號(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)</param>
    ''' <param name="newdata">儲存值</param>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    Private Sub Update_PlanTrainDescReviseItem(ByVal ptdriid As String, ByVal altdataitem As Integer, ByVal newdata As String, ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction)
        'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        ptdriid = TIMS.ClearSQM(ptdriid)
        If ptdriid = "" Then Return
        If ptdriid = "" Then
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= "#TC_05_001_chg,Update_PlanTrainDescReviseItem 傳入 ptdriid 為空 :" & vbCrLf
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'Call TIMS.WriteTraceLog(strErrmsg5)
            TIMS.LOG.Warn(strErrmsg5)
            Return
        End If
        Call TIMS.OpenDbConn(tmpTrans.Connection)
        Try
            Dim sqlStr As String = "UPDATE PLAN_TRAINDESC_REVISEITEM SET NewData=@newdata,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE() WHERE PTDRIID=@ptdriid"
            Dim UCmd As New SqlCommand(sqlStr, tmpConn, tmpTrans)
            With UCmd
                .Parameters.Clear()
                .Parameters.Add("newdata", SqlDbType.NVarChar).Value = newdata
                .Parameters.Add("ptdriid", SqlDbType.Int).Value = ptdriid
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .ExecuteNonQuery()
            End With
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            strErrmsg5 &= " #Update_PlanTrainDescReviseItem ; PTDRID:" & ptdriid
            strErrmsg5 &= " ; altdataitem:" & CStr(altdataitem)
            strErrmsg5 &= " ; newdata:" & newdata
            Call TIMS.WriteTraceLog(strErrmsg5)

            If flagDebugTest Then Throw ex
        End Try
    End Sub

    'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
    ''' <summary> INSERT INTO PLAN_TRAINDESC_REVISEITEM </summary>
    ''' <param name="ptdid">流水號</param>
    ''' <param name="ptdrid">PLAN_TRAINDESC_REVISE PTDRID</param>
    ''' <param name="iAltDataID"></param>
    ''' <param name="altdataitem">儲存詳資序號(1.STrainDate 2.PName 3.PHour 4.PTID 5.TechID 6.TechID2)</param>
    ''' <param name="olddata">儲存值-舊</param>
    ''' <param name="newdata">儲存值-新</param>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    Private Sub Insert_PlanTrainDescReviseItem(ByVal ptdid As Integer, ByVal ptdrid As Integer, ByVal iAltDataID As Integer, ByVal altdataitem As Integer,
                                               ByVal olddata As String, ByVal newdata As String, ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction)
        'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        If olddata = "" AndAlso newdata = "" Then Return '(新舊無資料不做新增)

        Call TIMS.OpenDbConn(tmpTrans.Connection)
        Try 'PK@PTDRIID
            Dim sqlS1 As String = "SELECT 1 FROM PLAN_TRAINDESC_REVISEITEM WHERE PTDRID=@ptdrid and PTDID=@ptdid and AltDataID=@AltDataID and AltDataItem=@AltDataItem"
            Dim sCmdS1 As New SqlCommand(sqlS1, tmpConn, tmpTrans)
            Dim dtS1 As New DataTable
            With sCmdS1
                .Parameters.Clear()
                .Parameters.Add("ptdrid", SqlDbType.Int).Value = ptdrid
                .Parameters.Add("ptdid", SqlDbType.Int).Value = ptdid
                .Parameters.Add("AltDataID", SqlDbType.Int).Value = iAltDataID
                .Parameters.Add("AltDataItem", SqlDbType.Int).Value = altdataitem
                dtS1.Load(.ExecuteReader())
            End With
            If dtS1.Rows.Count > 0 Then Return

            Dim sqlStr As String = String.Empty
            sqlStr &= " INSERT INTO PLAN_TRAINDESC_REVISEITEM (PTDRIID ,PTDRID ,PTDID ,AltDataID ,AltDataItem,OldData ,NewData,MODIFYACCT,MODIFYDATE)" & vbCrLf
            sqlStr &= " VALUES(@PTDRIID ,@ptdrid ,@ptdid ,@altdataid ,@altdataitem ,@olddata ,@newdata,@MODIFYACCT,GETDATE()) "
            Dim InsertCommand As New SqlCommand(sqlStr, tmpConn, tmpTrans)
            Dim iPTDRIID As Integer = DbAccess.GetNewId(tmpTrans, "PLAN_TRAINDESC_REVISEITEM_PTDR,PLAN_TRAINDESC_REVISEITEM,PTDRIID")
            With InsertCommand
                'PTDRIID 'PLAN_TRAINDESC_REVISEITEM_PTDR
                .Parameters.Clear()
                .Parameters.Add("PTDRIID", SqlDbType.Int).Value = iPTDRIID
                .Parameters.Add("ptdrid", SqlDbType.Int).Value = ptdrid
                .Parameters.Add("ptdid", SqlDbType.Int).Value = ptdid
                .Parameters.Add("altdataid", SqlDbType.Int).Value = iAltDataID
                .Parameters.Add("altdataitem", SqlDbType.Int).Value = altdataitem
                .Parameters.Add("olddata", SqlDbType.NVarChar).Value = olddata
                .Parameters.Add("newdata", SqlDbType.NVarChar).Value = newdata
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .ExecuteNonQuery()
            End With
            'sqlAdp.Dispose()
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)
            Common.MessageBox(Me, String.Concat(ex.Message, "-C4A8"))
            If flagDebugTest Then Throw ex
        End Try
    End Sub

    ''' <summary>INSERT PLAN_TRAINDESC_REVISE</summary>
    ''' <param name="iPlanid"></param>
    ''' <param name="sComidno"></param>
    ''' <param name="iSeqno"></param>
    ''' <param name="dCDate"></param>
    ''' <param name="iSubseqno"></param>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    ''' <returns></returns>
    Function INSERT_PLANTRAINDESCREVISE(ByVal iPlanid As Integer, ByVal sComidno As String, ByVal iSeqno As Integer, ByVal dCDate As DateTime, ByVal iSubseqno As Integer, ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction) As Integer
        Dim rst As Integer = 0 '(oPTDRID)
        Call TIMS.OpenDbConn(tmpTrans.Connection)
        'Dim sqlAdp As New SqlDataAdapter
        Dim sqlStr As String = String.Empty
        sqlStr = ""
        sqlStr &= " INSERT INTO PLAN_TRAINDESC_REVISE(PTDRID ,PlanID ,ComIDNO ,SeqNO ,CDate ,SubSeqNO,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sqlStr &= " VALUES(@PTDRID ,@planid ,@comidno ,@seqno ,@cdate ,@subseqno,@MODIFYACCT,GETDATE()) " 'select identity as NewID
        Dim iCmd As New SqlCommand(sqlStr, tmpConn, tmpTrans)
        Dim oPTDRID As Integer = DbAccess.GetNewId(tmpTrans, "PLAN_TRAINDESC_REVISE_PTDRID_S,PLAN_TRAINDESC_REVISE,PTDRID")
        rst = oPTDRID
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRID", SqlDbType.Int).Value = oPTDRID
            .Parameters.Add("planid", SqlDbType.Int).Value = iPlanid
            .Parameters.Add("comidno", SqlDbType.VarChar).Value = sComidno
            .Parameters.Add("seqno", SqlDbType.Int).Value = iSeqno
            .Parameters.Add("cdate", SqlDbType.DateTime).Value = dCDate
            .Parameters.Add("subseqno", SqlDbType.Int).Value = iSubseqno
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            .ExecuteNonQuery()
        End With
        '直接存入現在課表到 PLAN_TRAINDESC_RO
        Call INSERT_PLANTRAINDESCRO(oPTDRID, iPlanid, sComidno, iSeqno, tmpConn, tmpTrans)
        Return rst 'oPTDRID
    End Function

    ''' <summary>直接存入現在課表到 PLAN_TRAINDESC_RO</summary>
    ''' <param name="iPTDRID"></param>
    ''' <param name="iPlanid"></param>
    ''' <param name="sComidno"></param>
    ''' <param name="iSeqno"></param>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    Sub INSERT_PLANTRAINDESCRO(ByVal iPTDRID As Integer, ByVal iPlanid As Integer, ByVal sComidno As String, ByVal iSeqno As Integer, ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction)
        Dim sSql As String = ""
        sSql &= " INSERT INTO PLAN_TRAINDESC_RO (PTDRID,PTDID,PLANID,COMIDNO,SEQNO,STRAINDATE,ETRAINDATE,PNAME,PHOUR,EHOUR,PCONT,TRAINDEP" & vbCrLf
        sSql &= " ,CLASSIFICATION1,CLASSIFICATION2,CF1HOURS1,CF1HOURS2,CF2HOURS1,CF2HOURS2,PTID,TECHID,TECHID2,TPERIOD28,FARLEARN,OUTLEARN ,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sSql &= " SELECT @PTDRID ,PTDID,PLANID,COMIDNO,SEQNO,STRAINDATE,ETRAINDATE,PNAME,PHOUR,EHOUR,PCONT,TRAINDEP" & vbCrLf
        sSql &= " ,CLASSIFICATION1,CLASSIFICATION2,CF1HOURS1,CF1HOURS2,CF2HOURS1,CF2HOURS2,PTID,TECHID,TECHID2,TPERIOD28,FARLEARN,OUTLEARN ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sSql &= " FROM PLAN_TRAINDESC" & vbCrLf
        sSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        Dim iCmd As New SqlCommand(sSql, tmpConn, tmpTrans)
        'Dim oPTDRID As Integer=DbAccess.GetNewId(tmpTrans, "PLAN_TRAINDESC_REVISE_PTDRID_S,PLAN_TRAINDESC_REVISE,PTDRID")
        'rst=oPTDRID
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRID", SqlDbType.Int).Value = iPTDRID
            .Parameters.Add("PLANID", SqlDbType.Int).Value = iPlanid
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = sComidno
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = iSeqno
            .ExecuteNonQuery()
        End With
    End Sub

    '檢核資料
    Function Check_PlanTrainDescClassification(ByVal tmpID As Integer, ByVal tmpNO As String, ByVal tmpSNO As Integer,
                                               ByVal cf As String, ByVal oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        sql &= " SELECT 'x' FROM PLAN_TRAINDESC" & vbCrLf
        sql &= " WHERE PlanID=@planid AND ComIDNO=@comidno AND SeqNO=@seqno" & vbCrLf
        sql &= " AND Classification1=@cf "
        Dim sCmd As New SqlCommand(sql, oConn)
        TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("planid", SqlDbType.Int).Value = tmpID
            .Parameters.Add("comidno", SqlDbType.VarChar).Value = tmpNO
            .Parameters.Add("seqno", SqlDbType.Int).Value = tmpSNO
            .Parameters.Add("cf", SqlDbType.Char).Value = cf
            dt.Load(.ExecuteReader())
            'cnt=.SelectCommand.ExecuteScalar()
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    ''' <summary>'訓練場地 依地點代號 反推PTID KEY</summary>
    ''' <param name="PlaceID"></param>
    ''' <param name="ComIDNO"></param>
    ''' <returns></returns>
    Function Get_TrainPlacePTID(ByVal PlaceID As String, ByVal ComIDNO As String) As Integer
        'Dim sqlAdp As New SqlDataAdapter
        Dim rst As Integer = 0
        If PlaceID Is Nothing OrElse PlaceID = "" Then Return rst
        Dim oRst As Object
        Dim sqlStr As String = String.Empty
        sqlStr = " SELECT PTID FROM PLAN_TRAINPLACE WITH(NOLOCK) WHERE PlaceID=@PlaceID AND ComIDNO=@ComIDNO "
        Dim sCmd As New SqlCommand(sqlStr, gobjconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PlaceID", SqlDbType.VarChar).Value = PlaceID
            .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = ComIDNO
            oRst = .ExecuteScalar()
        End With
        If oRst IsNot Nothing Then rst = Val(oRst)
        Return rst
    End Function

    '(限定) 取得可用的下拉 上課地點
    Function Get_ListPTID(ByVal obj As ListControl, ByVal PlaceID_PTID1 As String, ByVal PlaceID_PTID2 As String, ByVal COMIDNO As String, ByVal i_type As Integer) As ListControl
        Dim sql As String = ""
        Select Case i_type
            Case 1 '計畫變更申請
                sql = " SELECT PTID ,PlaceNAME AS Name FROM PLAN_TRAINPLACE WHERE PlaceID IN ('" & PlaceID_PTID1 & "','" & PlaceID_PTID2 & "')" & " AND COMIDNO=@COMIDNO AND ModifyType IS NULL "
            Case 2 '變更結果顯示
                PlaceID_PTID1 = If(PlaceID_PTID1 <> "", PlaceID_PTID1, "0")
                PlaceID_PTID2 = If(PlaceID_PTID2 <> "", PlaceID_PTID2, "0")
                sql = " SELECT PTID ,PlaceNAME AS Name FROM PLAN_TRAINPLACE WHERE PTID IN ('" & PlaceID_PTID1 & "','" & PlaceID_PTID2 & "')" & " AND COMIDNO=@COMIDNO AND ModifyType IS NULL "
        End Select

        Dim parms As New Hashtable From {{"COMIDNO", COMIDNO}}
        Dim dt As DataTable = DbAccess.GetDataTable(sql, gobjconn, parms)
        With obj
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "PTID"
            .DataBind()
            If TypeOf obj Is DropDownList Then .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

    'create ViewState(vs_dtTaddress)
    Sub GetVSAddress()
        'ViewState(vs_dtTaddress)
        Dim dtSpace As New DataTable
        ViewState(vs_dtTaddress) = Nothing
        dtSpace.Columns.Add("PID")
        dtSpace.Columns.Add("PlaceID")
        dtSpace.Columns.Add("Name")
        dtSpace.Columns.Add("classification")
        dtSpace.Columns.Add("PTID")
        dtSpace.Columns.Add("PlaceNAME")
        For i As Integer = 0 To 4
            Dim dr As DataRow = dtSpace.NewRow()
            dr("PID") = i
            dr("PlaceID") = ""
            dr("Name") = If(i = 0, "======請選擇======", "")
            dr("classification") = ""
            dr("PTID") = ""
            dr("PlaceNAME") = ""
            dtSpace.Rows.Add(dr)
        Next
        ViewState(vs_dtTaddress) = dtSpace
    End Sub

    'iIDType 是指 將學數科場地做編號1,2,3,4; 'i_CIFICAT 指1學科或2術科
    'Function GetTaddressDDL(ByVal obj As ListControl, ByVal sPlaceID As String, ByVal iIDType As Integer, ByVal i_CIFICAT As Integer) As ListControl
    '    Dim tempdt As DataTable=Nothing
    '    Dim drArry() As DataRow=Nothing
    '    'Dim dr As DataRow
    '    'Dim drArry() As DataRow
    '    'Dim i As Integer
    '    tempdt=ViewState(vs_dtTaddress)
    '    tempdt.Rows(iIDType).Item(1)=""
    '    tempdt.Rows(iIDType).Item(2)=""
    '    tempdt.Rows(iIDType).Item(3)=""
    '    tempdt.Rows(iIDType).Item(4)=""
    '    tempdt.Rows(iIDType).Item(5)=""
    '    If sPlaceID <> "" Then
    '        Dim dr As DataRow=TIMS.Get_SciTechDR(rComIDNO, sPlaceID, i_CIFICAT, gobjconn)   '取得場地的地址
    '        If dr IsNot Nothing Then
    '            tempdt.Rows(iIDType).Item(1)=dr("PlaceID")
    '            tempdt.Rows(iIDType).Item(2)=dr("Name")
    '            tempdt.Rows(iIDType).Item(3)=dr("classification")
    '            tempdt.Rows(iIDType).Item(4)=dr("PTID")
    '            tempdt.Rows(iIDType).Item(5)=dr("PlaceNAME")
    '        End If
    '    End If
    '    'i_CIFICAT 指1學科或2術科
    '    Select Case i_CIFICAT
    '        Case 1 '學科
    '            drArry=tempdt.Select("PID IN (0,1,3) and Name <> ''")
    '        Case 2 '術科
    '            drArry=tempdt.Select("PID IN (0,2,4) and Name <> ''")
    '    End Select

    '    obj.Items.Clear()
    '    For i As Integer=0 To drArry.Length - 1
    '        obj.Items.Insert(i, New ListItem(drArry(i)("Name"), drArry(i)("PTID")))
    '    Next
    '    ViewState(vs_dtTaddress)=tempdt
    '    Return obj
    'End Function

#Region "NO USE"
    'Function GETvalue(ByVal ComIDNO As String, ByVal PlaceID As String, ByVal ClassType As Integer) As DataRow
    '    Dim dr As DataRow=Nothing
    '    Dim sql As String=""
    '    sql="" & vbCrLf
    '    sql &= " SELECT a.PlaceID ,a.PlaceNAME + ' (' + CONVERT(varchar, a.ZipCode) + '-' + CONVERT(varchar, a.ZIP6W) + ')'" & vbCrLf
    '    sql &= "  + c.CTName + b.ZipName + a.Address AS Name ,a.classification ,a.PTID ,a.PlaceNAME" & vbCrLf
    '    sql &= " FROM Plan_TrainPlace a" & vbCrLf
    '    sql &= " JOIN id_zip b ON a.ZipCode=b.ZipCode" & vbCrLf
    '    sql &= " JOIN id_city c ON b.CTID=c.CTID" & vbCrLf
    '    sql &= " WHERE 1=1" & vbCrLf
    '    If ClassType=1 Then     '學科
    '        sql &= " AND a.CLASSIFICATION IN (@CIFICAT,3) (a.classification=1 OR a.classification=3)" & vbCrLf
    '    ElseIf ClassType=2 Then '術科
    '        sql &= " AND (a.classification=2 OR a.classification=3)" & vbCrLf
    '    End If
    '    sql &= " AND a.ModifyType IS NULL AND a.ComIDNO='" & ComIDNO.ToString & "'" & vbCrLf
    '    sql &= " AND PlaceID='" & PlaceID & "'"
    '    dr=DbAccess.GetOneRow(sql, gobjconn)
    '    'If Not dr Is Nothing Then Return dr
    '    Return dr
    'End Function
#End Region

#End Region

    'Private Sub NewData14_1b_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewData14_1b.SelectedIndexChanged
    '    TaddressS2=GetTaddressDDL(TaddressS2, TIMS.GetListValue(NewData14_1b), 1, 1)
    '    '顯示必要資料列 (有AUTOPOSTBACK造成)
    '    'Call DisplayTR()
    '    'Page.RegisterStartupScript("myload1", "<script>ChangeState('" & sm.UserInfo.TPlanID & "' );</script>")
    'End Sub

    'Private Sub NewData14_2b_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewData14_2b.SelectedIndexChanged
    '    TaddressT2=GetTaddressDDL(TaddressT2, TIMS.GetListValue(NewData14_2b), 2, 2)
    '    '顯示必要資料列 (有AUTOPOSTBACK造成)
    '    'Call DisplayTR()
    '    'Page.RegisterStartupScript("myload1", "<script>ChangeState('" & sm.UserInfo.TPlanID & "' );</script>")
    'End Sub

    'Private Sub NewData14_3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewData14_3.SelectedIndexChanged
    '    TaddressS2=GetTaddressDDL(TaddressS2, TIMS.GetListValue(NewData14_3), 3, 1)
    '    '顯示必要資料列 (有AUTOPOSTBACK造成)
    '    'Call DisplayTR()
    '    'Page.RegisterStartupScript("myload1", "<script>ChangeState('" & sm.UserInfo.TPlanID & "' );</script>")
    'End Sub

    'Private Sub NewData14_4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewData14_4.SelectedIndexChanged
    '    TaddressT2=GetTaddressDDL(TaddressT2, TIMS.GetListValue(NewData14_4), 4, 2)
    '    '顯示必要資料列 (有AUTOPOSTBACK造成)
    '    'Call DisplayTR()
    '    'Page.RegisterStartupScript("myload1", "<script>ChangeState('" & sm.UserInfo.TPlanID & "' );</script>")
    'End Sub

    Private Sub DG_BusPackageNew_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_BusPackageNew.ItemCommand
        Const Cst_PKName As String = "BPID"
        'Dim objTable As HtmlTable=CType(DataGrid4Table, HtmlTable)
        Select Case e.CommandName
            Case "xedit"
                source.EditItemIndex = e.Item.ItemIndex
            Case "xdel"
                If Not Session("Revise_BusPackage") Is Nothing Then
                    Dim dt As DataTable
                    dt = Session("Revise_BusPackage")
                    If dt.Select(Cst_PKName & "='" & e.CommandArgument & "'").Length <> 0 Then
                        dt.Select(Cst_PKName & "='" & e.CommandArgument & "'")(0).Delete()
                    End If
                    Session("Revise_BusPackage") = dt
                    source.Visible = False
                    If dt.Rows.Count > 0 Then
                        source.Visible = True
                        source.DataSource = dt
                    End If
                End If
                source.EditItemIndex = -1
            Case "xsave"
                Dim dt As DataTable
                Dim dr As DataRow
                Dim tUName As TextBox = e.Item.FindControl("TUname")
                Dim tIntaxno As TextBox = e.Item.FindControl("TIntaxno")
                Dim tUbno As TextBox = e.Item.FindControl("TUbno")
                dt = Session("Revise_BusPackage")
                If dt.Select(Cst_PKName & "='" & e.CommandArgument & "'").Length <> 0 Then
                    dr = dt.Select(Cst_PKName & "='" & e.CommandArgument & "'")(0)
                    dr("Uname") = Convert.ToString(tUName.Text.Trim)
                    dr("Intaxno") = TIMS.ChangeIDNO(tIntaxno.Text)
                    dr("Ubno") = TIMS.ChangeIDNO(tUbno.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                Session("Revise_BusPackage") = dt
                source.EditItemIndex = -1
            Case "xcancel"
                source.EditItemIndex = -1
        End Select
        CreateBusPackage()
        'Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    Private Sub DG_BusPackageNew_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_BusPackageNew.ItemDataBound
        Const Cst_PKName As String = "BPID"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LUname As Label = e.Item.FindControl("LUname")
                Dim LIntaxno As Label = e.Item.FindControl("LIntaxno")
                Dim LUbno As Label = e.Item.FindControl("LUbno")
                Dim btnxEDT As Button = e.Item.FindControl("btnxEDT") '修改
                Dim btnxDEL As Button = e.Item.FindControl("btnxDEL") '刪除
                btnxEDT.Enabled = btnAddBusPackage.Enabled
                btnxDEL.Enabled = btnAddBusPackage.Enabled
                LUname.Text = drv("Uname").ToString
                LIntaxno.Text = drv("Intaxno").ToString
                LUbno.Text = drv("Ubno").ToString
                btnxEDT.CommandArgument = drv(Cst_PKName)
                btnxDEL.CommandArgument = drv(Cst_PKName)
                btnxDEL.Attributes("onclick") = TIMS.cst_confirm_delmsg1
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim TUname As TextBox = e.Item.FindControl("TUname")
                Dim TIntaxno As TextBox = e.Item.FindControl("TIntaxno")
                Dim TUbno As TextBox = e.Item.FindControl("TUbno")
                Dim btnxSAV As Button = e.Item.FindControl("btnxSAV") '儲存
                Dim btnxCLS As Button = e.Item.FindControl("btnxCLS") '取消
                TUname.Text = Convert.ToString(drv("Uname"))
                TIntaxno.Text = Convert.ToString(drv("Intaxno"))
                TUbno.Text = Convert.ToString(drv("Ubno"))
                btnxSAV.Enabled = btnAddBusPackage.Enabled
                btnxCLS.Enabled = btnAddBusPackage.Enabled
                btnxSAV.CommandArgument = drv(Cst_PKName)
                btnxCLS.CommandArgument = drv(Cst_PKName)
        End Select
    End Sub

    Private Sub btnAddBusPackage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddBusPackage.Click
        Const Cst_PKName As String = "BPID"
        'Plan_BusPackage
        ' Session("Revise_BusPackage")
        '錯誤檢查
        Dim Errmsg As String = ""
        If Trim(txtUname.Text) = "" Then
            txtUname.Text = ""
            Errmsg &= "企業名稱，不可為空" & vbCrLf
        Else
            '錯誤檢查
            txtUname.Text = Trim(txtUname.Text)
            If txtUname.Text.ToString.Length > 50 Then Errmsg &= "企業名稱，長度超過限制範圍50文字長度" & vbCrLf
        End If
        If Trim(txtIntaxno.Text) <> "" Then
            txtIntaxno.Text = Trim(txtIntaxno.Text)
            If Not TIMS.CheckIsECFA(TIMS.ChangeIDNO(txtIntaxno.Text), gobjconn) Then
                '未填寫 ECFA包班事業單位資料
                Errmsg &= "「" & Convert.ToString(txtUname.Text.Trim) & "」該企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf
            End If
        Else
            txtIntaxno.Text = ""
            Errmsg &= "服務單位統一編號，不可為空" & vbCrLf
        End If
        '錯誤檢查
        If Errmsg <> "" Then
            'Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        If Session("Revise_BusPackage") Is Nothing Then
            If rSCDate <> "" Then
                sql = ""
                sql &= " SELECT *" & vbCrLf
                sql &= " FROM Revise_BusPackage" & vbCrLf
                sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "'" & vbCrLf
                'sql &= "AND SCDate='" & rSCDate & "'" & vbCrLf
                sql &= " AND SCDate=" & TIMS.To_date(rSCDate) & vbCrLf
            Else
                sql = " SELECT * FROM Revise_BusPackage WHERE 1<>1 "
            End If
            dt = DbAccess.GetDataTable(sql, gobjconn)
            dt.Columns(Cst_PKName).AutoIncrement = True
            dt.Columns(Cst_PKName).AutoIncrementSeed = -1
            dt.Columns(Cst_PKName).AutoIncrementStep = -1
            Session("Revise_BusPackage") = dt
        Else
            dt = Session("Revise_BusPackage")
        End If
        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("PlanID") = rPlanID
        dr("ComIDNO") = rComIDNO
        dr("SeqNo") = rSeqNo
        dr("Uname") = Convert.ToString(txtUname.Text.Trim)
        dr("Intaxno") = TIMS.ChangeIDNO(txtIntaxno.Text)
        dr("Ubno") = TIMS.ChangeIDNO(txtUbno.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DG_BusPackageNew.Visible = True
        DG_BusPackageNew.DataSource = dt
        DG_BusPackageNew.DataBind()
        Session("Revise_BusPackage") = dt
        txtUname.Text = ""
        txtIntaxno.Text = ""
        txtUbno.Text = ""
        'Page.RegisterStartupScript("Londing", "<script>Layer_change(5);window.scroll(0,document.body.scrollHeight);</script>")
    End Sub

    Public Shared Sub SAVE_REVISE_BUSPACKAGE(ByRef MyPage As Page, ByRef htSS As Hashtable, ByRef oConn As SqlConnection, ByRef oTrans As SqlTransaction)
        Dim sql As String = TIMS.GetMyValue2(htSS, "sql")
        Dim rPlanID As Integer = TIMS.GetMyValue2(htSS, "rPlanID")
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO")
        Dim rSeqNo As Integer = TIMS.GetMyValue2(htSS, "rSeqNo")
        Dim iSubSeqNO As Integer = TIMS.GetMyValue2(htSS, "iSubSeqNO")
        Dim sApplyDate As String = TIMS.GetMyValue2(htSS, "sApplyDate")
        Dim sAltDataID As String = TIMS.GetMyValue2(htSS, "sAltDataID")
        Dim iPlanKind As Integer = TIMS.GetMyValue2(htSS, "iPlanKind")
        Dim OldData4_1 As String = TIMS.GetMyValue2(htSS, "OldData4_1")
        Dim NewData4_1 As String = TIMS.GetMyValue2(htSS, "NewData4_1")
        Dim sReviseCont As String = TIMS.GetMyValue2(htSS, "sReviseCont")
        Dim v_changeReason As String = TIMS.GetMyValue2(htSS, "changeReason")
        Dim sPackageTypeNew As String = TIMS.GetMyValue2(htSS, "sPackageTypeNew")
        Dim txtUname As String = TIMS.GetMyValue2(htSS, "txtUname")
        Dim txtIntaxno As String = TIMS.GetMyValue2(htSS, "txtIntaxno")
        Dim txtUbno As String = TIMS.GetMyValue2(htSS, "txtUbno")
        Dim rPARTREDUC1 As String = TIMS.GetMyValue2(htSS, "rPARTREDUC1")
        Dim sm As SessionModel = SessionModel.Instance()

        'Dim oTrans As SqlTransaction=DbAccess.BeginTrans(gobjconn)
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = DbAccess.GetDataTable(sql, da, oTrans)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0) '1天提供1筆
        Else
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("PlanID") = rPlanID 'Request("PlanID")
            dr("ComIDNO") = rComIDNO 'Request("cid")
            dr("SeqNO") = rSeqNo 'Request("no")
            dr("SubSeqNo") = iSubSeqNO
            dr("CDate") = CDate(sApplyDate) '.ToString("yyyy/MM/dd")
            dr("AltDataID") = sAltDataID 'ChgItem.SelectedValue
        End If
        If rPARTREDUC1 <> "" Then
            'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕 'Dim flag_PARTREDUC_Y_CanUpdate As Boolean=False 'If rPARTREDUC1="Y" AndAlso Convert.ToString(drPR("PARTREDUC"))="Y" AndAlso Convert.ToString(drPR("ReviseStatus"))="" Then flag_PARTREDUC_Y_CanUpdate=True
            Dim flag_PARTREDUC_Y_CanUpdate As Boolean = (rPARTREDUC1 = "Y" AndAlso $"{dr("PARTREDUC")}" = "Y" AndAlso $"{dr("ReviseStatus")}" = "")
            If flag_PARTREDUC_Y_CanUpdate Then dr("PARTREDUC") = Convert.DBNull
        End If
        dr("OldData4_1") = OldData4_1 'hidPackageTypeOld.Value'Cst_i包班種類
        dr("NewData4_1") = NewData4_1 'PackageTypeNew.SelectedValue'Cst_i包班種類
        '變更-審核人員
        dr("ReviseAcct") = sm.UserInfo.UserID
        '變更內容說明
        dr("ReviseCont") = If(sReviseCont <> "", sReviseCont, Convert.DBNull)
        '變更原因說明
        dr("changeReason") = If(v_changeReason <> "", v_changeReason, Convert.DBNull)
        dr("Verifier") = If(iPlanKind = 1, sm.UserInfo.UserID, Convert.DBNull)
        dr("ReviseStatus") = If(iPlanKind = 1, "Y", Convert.DBNull)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()
        DbAccess.UpdateDataTable(dt, da, oTrans)

        sql = ""
        sql &= " DELETE REVISE_BUSPACKAGE" & vbCrLf
        sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNo & "' AND SCDate=" & TIMS.To_date(sApplyDate) & vbCrLf
        DbAccess.ExecuteNonQuery(sql, oTrans)

        Select Case sPackageTypeNew 'PackageTypeNew.SelectedValue
            Case "2" '充電起飛計畫 '企業包班
                Dim iBPID As Integer = DbAccess.GetNewId(oTrans, "REVISE_BUSPACKAGE_BPID_SEQ,REVISE_BUSPACKAGE,BPID")
                sql = " SELECT * FROM REVISE_BUSPACKAGE WHERE 1<>1 "
                dt = DbAccess.GetDataTable(sql, da, oTrans)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                'REVISE_BUSPACKAGE_BPID_SEQ
                dr("BPID") = iBPID
                dr("PlanID") = rPlanID 'Request("PlanID")
                dr("ComIDNO") = rComIDNO 'Request("cid")
                dr("SeqNO") = rSeqNo 'Request("no")
                'dr("SubSeqNo")=iSubSeqNO
                dr("SCDate") = CDate(sApplyDate).ToString("yyyy/MM/dd")
                dr("UName") = "" & txtUname '.Text.Trim
                dr("Intaxno") = "" & txtIntaxno '.Text.Trim
                dr("Ubno") = "" & txtUbno '.Text.Trim
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now()
                DbAccess.UpdateDataTable(dt, da, oTrans)
            Case "3" '充電起飛計畫 '聯合企業包班
                '申請只會送一次
                Dim dtTemp As DataTable = Nothing
                dtTemp = MyPage.Session("Revise_BusPackage")
                sql = " SELECT * FROM REVISE_BUSPACKAGE WHERE 1<>1 "
                dt = DbAccess.GetDataTable(sql, da, oTrans)
                For Each drT As DataRow In dtTemp.Rows
                    If Not drT.RowState = DataRowState.Deleted Then
                        Dim iBPID As Integer = DbAccess.GetNewId(oTrans, "REVISE_BUSPACKAGE_BPID_SEQ,REVISE_BUSPACKAGE,BPID")
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("BPID") = iBPID
                        dr("PlanID") = rPlanID 'Request("PlanID")
                        dr("ComIDNO") = rComIDNO 'Request("cid")
                        dr("SeqNO") = rSeqNo 'Request("no")
                        'dr("SubSeqNo")=iSubSeqNO
                        dr("SCDate") = CDate(sApplyDate).ToString("yyyy/MM/dd")
                        dr("UName") = drT("UName")
                        dr("Intaxno") = drT("Intaxno")
                        dr("Ubno") = drT("Ubno")
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                    End If
                Next

                'dt=dtTemp.Copy
                DbAccess.UpdateDataTable(dt, da, oTrans)
        End Select
    End Sub

#Region "NO USE"

    '回上一頁
    'Private Sub back_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles back.PreRender
    '    ViewState(vs_Do_CreateTrainDesc)="N"
    'End Sub

    '回上一頁
    'Private Sub back_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles back.ServerClick
    '    'Session("_search")=ViewState("_search")
    '    ViewState(vs_TEMP11_TrainDescDT)=""
    '    ViewState(vs_TEMP20_TrainDescDT)=""
    '    'Response.Redirect("TC_05_001.aspx?ID=" & Request("ID"))
    '    Dim url1 As String="TC_05_001.aspx?ID=" & Request("ID")
    '    Call TIMS.Utl_Redirect(Me, gobjconn, url1)
    'End Sub

    'select * from CLASS_CLASSINFO where ocid =46620
    'select 'planid='+convert(varchar,planid)+'comidno='+comidno+'seqno='+convert(varchar,seqno)  xx
    'from CLASS_CLASSINFO where ocid =46620
    'select * from Plan_BusPackage a
    'where 1=1
    'AND a.PlanID='1783' and a.ComIDNO='94991877'  
    'select * from Revise_BusPackage a
    'where 1=1
    'AND a.PlanID='1783' and a.ComIDNO='94991877'  
    'select * from PLAN_TRAINDESC a 
    'where 1=1
    '   AND a.PlanID='1783' and a.ComIDNO='94991877' --and a.SeqNo='2' 
    'select * from PLAN_TRAINDESC a 
    'where 1=1
    '   AND a.PlanID='1783'
    'select * from REVISE_ONCLASS a
    'where 1=1
    '   AND a.PlanID='1783' and a.ComIDNO='94991877' -- and a.SeqNo='2' 
    'select * from REVISE_ONCLASS a
    'where 1=1
    '   AND a.PlanID='1783'  
    'select *
    'from Plan_OnClass a
    'where 1=1
    '   AND a.PlanID='1783' and a.ComIDNO='94991877' -- and a.SeqNo='2' 
    'select b.* 
    'from Teach_TeacherInfo b
    'join PLAN_TRAINDESC a on a.techid2=b.techid
    'where 1=1
    '   AND a.PlanID='1783' and a.ComIDNO='94991877' -- and a.SeqNo='2' 
    'select b.* 
    'from Teach_TeacherInfo b
    'join PLAN_TRAINDESC a on a.techid=b.techid
    'where 1=1
    '   AND a.PlanID='1783' and a.ComIDNO='94991877' -- and a.SeqNo='2' 
    'select sc.*
    'from CLASS_SCHEDULE sc
    'join CLASS_CLASSINFO cc on sc.ocid =cc.ocid
    'join PLAN_PLANINFO pp on pp.planid =cc.planid and pp.comidno =cc.comidno and pp.seqno =cc.seqno 
    'where 1=1
    '   AND pp.PlanID='1783' and pp.ComIDNO='94991877' -- and pp.SeqNo='2' 
    'select top 10 * from PLAN_TRAINDESC_REVISEITEM
    'select top 10 * from PLAN_TRAINDESC_REVISE
    'select * 
    'from PLAN_TRAINDESC_REVISE pp
    'where 1=1
    '   AND pp.PlanID='1783' --and pp.ComIDNO='94991877' -- and pp.SeqNo='2' 
    'select p2.* 
    'from PLAN_TRAINDESC_REVISE pp
    'join PLAN_TRAINDESC_REVISEITEM p2 on p2.PTDRID=pp.PTDRID
    'where 1=1
    '   AND pp.PlanID='1783' --and pp.ComIDNO='94991877' -- and pp.SeqNo='2' 
    'select * 
    'from Plan_TrainPlace pp 
    'where 1=1 and pp.ComIDNO='94991877'

    'REVISE_ONCLASS 

    'Dim j As Integer
    'Dim table As DataTable
    'Dim chgName As Array    '儲存ChgItem.Items.Text字串的陣列
    'Dim firstLoad As Boolean=False   '用來判斷是否初始化過

    '? Request.RawUrl
    'select *
    'FROM PLAN_PLANINFO a
    'where 1=1
    'and a.PlanID=1555 and a.ComIDNO=17154491---  and a.SeqNO=1
    'SELECT * FROM PLAN_TRAINDESC a
    'where 1=1
    'and a.PlanID=1555 and a.ComIDNO=17154491---  and a.SeqNO=1
    'SELECT * FROM Plan_CostItem
    'where 1=1
    'and PlanID=1555 and ComIDNO=17154491---  and SeqNO=1
    'select *
    'FROM CLASS_CLASSINFO a
    'where 1=1
    'and a.PlanID=1555 and a.ComIDNO=17154491---  and a.SeqNO=1
    'select * from View_CLASS_SCHEDULE where ocid =41123
    'SELECT * FROM CLASS_SCHEDULE  where ocid =41123
    'SELECT * FROM Class_StudentsOfClass where ocid =41123
    'select a.*
    'FROM Class_Teacher a 
    'JOIN Teach_TeacherInfo b ON a.TechID=b.TechID
    'where exists (
    '	select 'x'
    '	FROM CLASS_CLASSINFO x
    '	where 1=1
    '	and x.PlanID=1555 and x.ComIDNO=17154491---  and a.SeqNO=1
    '	and x.ocid =a.ocid 
    ')
    'select top 18 * from Class_Teacher
    'select top 18 * from Teach_TeacherInfo
    '---- select top 10 * from view_ClassinfoSch where ocid =41123 
    'select * from org_orginfo where ComIDNO=17154491
    'select * from org_orginfo where orgid=619 and ComIDNO=17154491
    'select *  from Course_CourseInfo  where orgid=619
    'SELECT * 
    'FROM Teach_TeacherInfo P
    'WHERE EXISTS  (
    '	select 'x'  from Course_CourseInfo  x where x.orgid=619
    '	and x.RID =p.RID
    ')

    '20081127 "課程大綱"改成"課程表"
    '**by Milor 20080507--將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改----start
    'If firstLoad Then
    '    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then '產學訓套用的顯示字串，但會有年度之分
    '        If sm.UserInfo.Years <= 2007 Then    '產學訓96年之前的顯示字串
    '            chgName=New String() {"訓練期間", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別", "期別", "上課地址", "申請停辦", "上課時段", "師資", "招生人數", "增班", "學(術)科場地", "上課時間", "其他", "報名日期", "課程表"}
    '        Else    '產學訓97年以後的顯示字串
    '            chgName=New String() {"開、結訓日期", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別", "期別", "上課地址", "停辦", "上課時段", "師資", "招生人數", "增班", "上課地點", "上課時間", "其他", "報名日期", "課程表"}
    '        End If
    '    Else    '非產學訓套用的顯示字串
    '        '    chgName=New String() {"訓練期間", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別名稱", "期別", "上課地址", "申請停辦", "上課時段", "師資", "招生人數", "增班", "學(術)科場地", "上課時間", "其他", "課程表"}
    '        chgName=New String() {"訓練期間", "訓練時段", "訓練課程地點", "課程編配", "訓練師資", "班別名稱", "期別", "上課地址", "申請停辦", "上課時段", "師資", "招生人數", "增班", "學(術)科場地", "上課時間", "其他", "報名日期"}
    '    End If
    '    For j=1 To ChgItem.Items.Count - 1    '將顯示名稱套用
    '        ChgItem.Items(j).Text=chgName(j - 1)
    '    Next
    '    firstLoad=True
    'End If

    'Function GetAltStr(ByVal dateText As TextBox, ByVal lb As ListBox, ByVal TypeStr As String)
    '    Dim ReturnStr As String=""
    '    Dim n, x, i As Integer
    '    Dim sql As String
    '    Dim dr As DataRow
    '    Dim SelTypeStr As String
    '    Dim lbarray As New ArrayList
    '    For i=0 To lb.Items.Count - 1
    '        lbarray.Add(lb.Items(i).Value)
    '    Next
    '    Dim result As String
    '    For i=0 To lbarray.Count - 1
    '        sql="select " & TypeStr & Convert.ToString(lbarray(i)) & " from CLASS_SCHEDULE" & _
    '              " where OCID=" & ViewState(vs_OCID) & " and SchoolDate ='" & dateText.Text & "'"
    '        If Convert.ToString(DbAccess.ExecuteScalar(sql)) <> "" Then
    '            SelTypeStr=SelTypeStr & Convert.ToString(DbAccess.ExecuteScalar(sql)) & ","
    '        Else
    '            SelTypeStr=SelTypeStr & "x" & ","
    '        End If
    '    Next
    '    ReturnStr=Left(SelTypeStr, Len(SelTypeStr) - 1)
    '    Return ReturnStr
    'End Function

    ''20080902 andy  訓練時段 確認選擇日期當天是否有排課
    'Function ChkSchoolDate(ByVal SchoolDate As TextBox, ByVal msg As Label)
    '    Dim sql As String
    '    Dim dr As DataRow
    '    Dim i As Integer
    '    Dim j As Integer
    '    'Dim dt As DataTable

    '    'sql=" select  ocid   from CLASS_SCHEDULE where  OCID=" & ViewState(vs_OCID)
    '    'dt=DbAccess.GetDataTable(sql)
    '    'If dt.Rows.Count=0 Then
    '    '    If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '    '        If hid_chkmsg.Value="on" Then
    '    '            hid_chkmsg.Value="off"
    '    '            Common.MessageBox(Page, "本班目前尚未排課，請先確認是否已於課程管理完成排課作業！")
    '    '        End If
    '    '    End If
    '    '    Exit Function
    '    'End If

    '    Dim sErrmsg As String=""
    '    If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        If Not CheckClassSchedule(ViewState(vs_OCID), sErrmsg) Then
    '            If hid_chkmsg.Value="on" Then
    '                msg.Text=sErrmsg
    '                hid_chkmsg.Value="off"
    '                Common.MessageBox(Page, sErrmsg)
    '            End If
    '            Exit Function
    '        End If
    '    End If

    '    If ViewState(vs_OCID) <> "" Then
    '        sql="select Teacher1,Teacher2,Teacher3,Teacher4,Teacher5,Teacher6,Teacher7,Teacher8,Teacher9,Teacher10,Teacher11,Teacher12," & _
    '             " Class1,Class2,Class3,Class4,Class5,Class6,Class7,Class8,Class9,Class10,Class11,Class12" & _
    '             " from CLASS_SCHEDULE" & _
    '             " where OCID=" & ViewState(vs_OCID) & " and SchoolDate='" & SchoolDate.Text & "'"
    '        dr=DbAccess.GetOneRow(sql)

    '        If dr Is Nothing Then
    '            msg.Text="查無資料(該班的開結訓日期範圍" & TRange.Text & ")"
    '        Else
    '            j=0
    '            For i=1 To 12
    '                If IsDBNull(dr("Class" & i)) Then
    '                    j += 1
    '                End If
    '            Next
    '            If j=12 Then
    '                msg.Text="當天無課程資料"
    '            End If
    '        End If

    '    End If
    'End Function

    'Private Sub bt_clearTech_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_clearTech.Click
    '    OLessonTeah1Value.Value=""
    '    OLessonTeah2Value.Value=""
    '    OLessonTeah1.Text=""
    '    OLessonTeah2.Text=""
    'End Sub

    'Function IsInt(ByVal chkstr As String) As Boolean  '20090318 andy add 判斷是否為正整數
    '    Dim dnum As Double=0
    '    If IsNumeric(chkstr) Then
    '        dnum=chkstr
    '        If CInt(dnum).ToString=Trim(dnum.ToString) Then
    '            Return True
    '        Else
    '            Return False
    '        End If
    '    Else
    '        Return False
    '    End If
    'End Function

    'Private Function Get_ClassRow(ByVal dateText As TextBox, ByVal da As SqlDataAdapter) As DataRow
    '    da=TIMS.GetOneDA()
    '    Dim dt As New DataTable
    '    Dim sql As String
    '    sql="select  *   from CLASS_SCHEDULE" & _
    '          " where OCID=" & ViewState(vs_OCID) & " and SchoolDate ='" & dateText.Text & "'"
    '    da.SelectCommand.CommandText=sql
    '    da.Fill(dt)
    '    If dt.Rows.Count=0 Then
    '        dt.Clear()
    '        sql="select  *   from CLASS_SCHEDULE  where  1<>1 "
    '        da.SelectCommand.CommandText=sql
    '        da.Fill(dt)
    '    End If
    '    Return dt.Rows(0)
    'End Function

    ''判斷是數字且小於20
    'Function sUtl_xNum(ByVal str As String) As Boolean
    '    Dim rst As Boolean=False
    '    'AndAlso sUtl_xNum(SPlace.Items(i).Value) 
    '    If IsNumeric(str) Then
    '        If CInt(str) < 20 Then
    '            rst=True
    '        End If
    '    End If
    '    Return rst
    'End Function

    'Public Sub SetClassSchedule()
    '    Dim dt1 As DataTable=Nothing
    '    Dim dt2 As DataTable=Nothing
    '    Dim dr1 As DataRow=Nothing
    '    Dim dr2 As DataRow=Nothing
    '    Dim da1 As SqlDataAdapter=Nothing
    '    Dim da2 As SqlDataAdapter=Nothing
    '    Dim dr As DataRow=Nothing

    '    Dim CourseTable As DataTable=Nothing  '課程資料
    '    Dim Holiday As DataTable=Nothing '假別資料
    '    Dim conn As SqlConnection=DbAccess.GetConnection()

    '    Dim sql As String=""
    '    sql="SELECT a.RID,c.OrgID,d.OCID FROM "
    '    sql += "(SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "') a "
    '    sql += "JOIN Auth_Relship b ON a.RID=b.RID "
    '    sql += "JOIN Org_OrgInfo c ON b.OrgID=c.OrgID "
    '    sql += "JOIN (SELECT * FROM CLASS_CLASSINFO WHERE PlanID='" & rPlanID & "' and ComIDNO='" & rComIDNO & "' and SeqNO='" & rSeqNo & "') d ON a.PlanID=d.PlanID and a.ComIDNO=d.ComIDNO and a.SeqNo=d.SeqNO "
    '    sql += "JOIN Class_TmpSchedule e ON e.OCID=d.OCID "
    '    dr=DbAccess.GetOneRow(sql, conn)

    '    If Not dr Is Nothing Then
    '        Dim OrgID As Integer=dr("OrgID")
    '        Dim RID As String=dr("RID")
    '        Dim OCID As String=dr("OCID")

    '        sql="SELECT * FROM Course_CourseInfo WHERE OrgID='" & OrgID & "'"
    '        CourseTable=DbAccess.GetDataTable(sql, conn)
    '        sql="SELECT * FROM Sys_Holiday WHERE RID='" & RID & "'"
    '        Holiday=DbAccess.GetDataTable(sql, conn)

    '        '先刪除已經排入課程的資料
    '        sql="DELETE CLASS_SCHEDULE WHERE OCID='" & OCID & "'"
    '        DbAccess.ExecuteNonQuery(sql, conn)

    '        sql="SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & OCID & "'"
    '        dr=DbAccess.GetOneRow(sql, conn)
    '        Dim STDate As Date=dr("STDate")
    '        Dim FTDate As Date=dr("FTDate")
    '        Dim TotalHour As Integer=dr("THours")

    '        sql="SELECT * FROM Class_TmpSchedule WHERE OCID='" & OCID & "' Order By ItemID"
    '        dt1=DbAccess.GetDataTable(sql, da1, conn)
    '        sql="SELECT * FROM CLASS_SCHEDULE WHERE 1<>1"
    '        dt2=DbAccess.GetDataTable(sql, da2, conn)

    '        Dim TempDate As Date=STDate
    '        While TempDate <= FTDate
    '            dr2=dt2.NewRow
    '            dt2.Rows.Add(dr2)
    '            dr2("OCID")=OCID
    '            dr2("SchoolDate")=TempDate
    '            dr2("Formal")="N"
    '            dr2("Type")=1
    '            dr2("ModifyAcct")=sm.UserInfo.UserID
    '            dr2("ModifyDate")=Now

    '            TempDate=TempDate.AddDays(1)
    '        End While

    '        Dim RealHours As Integer
    '        Dim CalHours As Integer
    '        Dim WeekIndex As Integer
    '        For Each dr1 In dt1.Rows
    '            RealHours=0
    '            CalHours=dr1("CalHours")
    '            WeekIndex=1
    '            For Each dr2 In dt2.Select("SchoolDate>='" & dr1("StartDate") & "' and SchoolDate<='" & dr1("EndDate") & "'", "SchoolDate")
    '                If Holiday.Select("HolDate='" & dr2("SchoolDate") & "'").Length=0 Then            '非假日判斷
    '                    '循環判斷
    '                    Dim Recycle As Boolean=True
    '                    If dr1("Recycle").ToString <> "" Then
    '                        If WeekIndex Mod Int(dr1("Recycle")) <> 1 And Int(dr1("Recycle")) <> 1 Then
    '                            Recycle=False
    '                        End If
    '                    End If
    '                    If Recycle=True Then
    '                        For i As Integer=1 To 7
    '                            If Not IsDBNull(dr1("S" & i)) And (Weekday(dr2("SchoolDate"))=i + 1 Or Weekday(dr2("SchoolDate"))=i - 6) Then
    '                                'Dim j As Integer
    '                                'Dim k As Integer
    '                                For j As Integer=dr1("S" & i) To dr1("E" & i)
    '                                    If IsDBNull(dr2("Class" & j)) Then
    '                                        If TotalHour > 0 And CalHours <> 0 Then
    '                                            dr2("Class" & j)=dr1("CourseID")
    '                                            dr2("Teacher" & j)=dr1("LessonTeah1")
    '                                            dr2("Teacher" & j + 12)=dr1("LessonTeah2")
    '                                            dr2("Room" & j)=dr1("RoomID")

    '                                            RealHours += 1
    '                                            CalHours -= 1
    '                                            TotalHour -= 1
    '                                        End If
    '                                    End If
    '                                Next
    '                            End If
    '                        Next
    '                    End If
    '                End If
    '                If Weekday(dr2("SchoolDate"))=1 Then
    '                    WeekIndex += 1
    '                End If
    '            Next

    '            dr1("RealHours")=RealHours
    '        Next

    '        DbAccess.UpdateDataTable(dt1, da1)
    '        DbAccess.UpdateDataTable(dt2, da2)
    '    End If

    '    Call TIMS.CloseDbConn(conn)
    'End Sub

    'Private Function GetClassName(ByVal ClassID)
    '    If IsDBNull(ClassID) Then
    '        Exit Function
    '    Else
    '        If Trim(ClassID)="" Then
    '            Exit Function
    '        End If
    '    End If
    '    Dim str As String="select CourseName from Course_CourseInfo where CourID in  (" & ClassID & ")"
    '    Return DbAccess.ExecuteScalar(str)
    'End Function

    'Private Function Get_TeacherName(ByVal TeacherID)
    '    If IsDBNull(TeacherID) Then
    '        Exit Function
    '    Else
    '        If Trim(TeacherID)="" Then
    '            Exit Function
    '        End If
    '    End If
    '    Dim str As String="select TeachCName from Teach_TeacherInfo where TechID  in (" & TeacherID & ")"
    '    Return DbAccess.ExecuteScalar(str)
    'End Function

    'Private Sub NewData14_1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If NewData14_1.SelectedIndex <> 0 Then
    '        TaddressS2=GetTaddresstable(TaddressS2, NewData14_1.SelectedValue, 1, 1)
    '    Else
    '        TaddressS2=GetTaddresstable(TaddressS2, "", 1, 1)
    '    End If
    'End Sub

    'Private Sub NewData14_2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If NewData14_2.SelectedIndex <> 0 Then
    '        TaddressT2=GetTaddresstable(TaddressT2, NewData14_2.SelectedValue, 2, 2)
    '    Else
    '        TaddressT2=GetTaddresstable(TaddressT2, "", 2, 2)
    '    End If
    'End Sub

    'ChgItem.Attributes("onchange")="ChangeState('" & sm.UserInfo.TPlanID & "');"
    'Protected Sub ChgItem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ChgItem.SelectedIndexChanged
    'End Sub

#End Region

    'UPDATE REVISE_COSTITEM / REVISE_COSTITEM_L1
    Public Shared Sub SAVE_REVISE_COSTITEM(ByVal htSS As Hashtable, ByRef dtTemp2 As DataTable, ByRef Conn As SqlConnection, ByRef Trans As SqlTransaction)
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        Dim rApplyDate As String = TIMS.GetMyValue2(htSS, "rApplyDate") 'ApplyDate.Text
        Dim iSubSeqNO As Integer = TIMS.CINT1(TIMS.GetMyValue2(htSS, "SubSeqNO")) 'iSubSeqNO
        Dim iCostMode As Integer = TIMS.CINT1(TIMS.GetMyValue2(htSS, "CostMode")) 'CostMode
        Dim ssUserID As String = TIMS.GetMyValue2(htSS, "ssUserID")

        Dim sql As String = ""
        sql = " SELECT * FROM PLAN_COSTITEM WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO ORDER BY PCID "
        Dim sCmdC1 As New SqlCommand(sql, Conn, Trans)

        sql = "" & vbCrLf
        sql &= " INSERT INTO REVISE_COSTITEM (RCID ,O1N2 ,COSTMODE ,PLANID ,COMIDNO ,SEQNO ,CDATE ,SUBSEQNO ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        sql &= " VALUES (@RCID ,@O1N2 ,@COSTMODE ,@PLANID ,@COMIDNO ,@SEQNO ,@CDATE ,@SUBSEQNO ,@MODIFYACCT ,GETDATE())" & vbCrLf
        Dim iCmdC1 As New SqlCommand(sql, Conn, Trans)

        sql = "" & vbCrLf
        sql &= " INSERT INTO REVISE_COSTITEM_L1(RCNID ,RCID ,O1N2 ,COSTID ,ITEMOTHER ,OPRICE ,ITEMAGE ,ITEMCOST ,ADMFLAG ,TAXFLAG ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        sql &= " VALUES (@RCNID ,@RCID ,@O1N2 ,@COSTID ,@ITEMOTHER ,@OPRICE ,@ITEMAGE ,@ITEMCOST ,@ADMFLAG ,@TAXFLAG ,@MODIFYACCT ,GETDATE())" & vbCrLf
        Dim iCmdC2 As New SqlCommand(sql, Conn, Trans)

        'PLAN_COSTITEM
        Dim dtTemp1 As New DataTable
        With sCmdC1
            .Parameters.Clear()
            .Parameters.Add("PLANID", SqlDbType.Int).Value = rPlanID
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = rSeqNo
            dtTemp1.Load(.ExecuteReader())
        End With
        Dim iRCID As Integer = 0
        iRCID = DbAccess.GetNewId(Trans, "REVISE_COSTITEM_RCID_SEQ,REVISE_COSTITEM,RCID")
        With iCmdC1
            .Parameters.Clear()
            .Parameters.Add("RCID", SqlDbType.Int).Value = iRCID
            .Parameters.Add("O1N2", SqlDbType.Int).Value = 1
            .Parameters.Add("COSTMODE", SqlDbType.Int).Value = iCostMode
            .Parameters.Add("PLANID", SqlDbType.Int).Value = rPlanID
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = rSeqNo
            .Parameters.Add("CDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(rApplyDate)
            .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = iSubSeqNO
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID 'Convert.ToString(sm.UserInfo.UserID)
            .ExecuteNonQuery()
        End With
        For Each drV As DataRow In dtTemp1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
            Dim iRCNID As Integer = DbAccess.GetNewId(Trans, "REVISE_COSTITEM_L1_RCNID_SEQ,REVISE_COSTITEM_L1,RCNID")
            With iCmdC2
                .Parameters.Clear()
                .Parameters.Add("RCNID", SqlDbType.Int).Value = iRCNID
                .Parameters.Add("RCID", SqlDbType.Int).Value = iRCID
                .Parameters.Add("O1N2", SqlDbType.Int).Value = 1
                .Parameters.Add("COSTID", SqlDbType.VarChar).Value = drV("COSTID")
                .Parameters.Add("ITEMOTHER", SqlDbType.VarChar).Value = drV("ITEMOTHER")
                .Parameters.Add("OPRICE", SqlDbType.Float).Value = drV("OPRICE")
                .Parameters.Add("ITEMAGE", SqlDbType.Int).Value = drV("ITEMAGE")
                .Parameters.Add("ITEMCOST", SqlDbType.Int).Value = drV("ITEMCOST")
                .Parameters.Add("ADMFLAG", SqlDbType.VarChar).Value = drV("ADMFLAG")
                .Parameters.Add("TAXFLAG", SqlDbType.VarChar).Value = drV("TAXFLAG")
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID 'Convert.ToString(sm.UserInfo.UserID)
                .ExecuteNonQuery()
            End With
        Next

        'Dim iRCID As Integer=0
        iRCID = DbAccess.GetNewId(Trans, "REVISE_COSTITEM_RCID_SEQ,REVISE_COSTITEM,RCID")
        With iCmdC1
            .Parameters.Clear()
            .Parameters.Add("RCID", SqlDbType.Int).Value = iRCID
            .Parameters.Add("O1N2", SqlDbType.Int).Value = 2
            .Parameters.Add("COSTMODE", SqlDbType.Int).Value = iCostMode 'Val(Hid_CostMode.Value)
            .Parameters.Add("PLANID", SqlDbType.Int).Value = rPlanID
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = rSeqNo
            .Parameters.Add("CDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(rApplyDate) 'ApplyDate.Text)
            .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = iSubSeqNO
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID 'Convert.ToString(sm.UserInfo.UserID)
            .ExecuteNonQuery()
        End With
        'Dim dtTemp2 As DataTable=Session(Hid_COSTITEM_GUID21.Value)
        For Each drV As DataRow In dtTemp2.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
            Dim iRCNID As Integer = DbAccess.GetNewId(Trans, "REVISE_COSTITEM_L1_RCNID_SEQ,REVISE_COSTITEM_L1,RCNID")
            With iCmdC2
                .Parameters.Clear()
                .Parameters.Add("RCNID", SqlDbType.Int).Value = iRCNID
                .Parameters.Add("RCID", SqlDbType.Int).Value = iRCID
                .Parameters.Add("O1N2", SqlDbType.Int).Value = 2
                .Parameters.Add("COSTID", SqlDbType.VarChar).Value = drV("COSTID")
                .Parameters.Add("ITEMOTHER", SqlDbType.VarChar).Value = drV("ITEMOTHER")
                .Parameters.Add("OPRICE", SqlDbType.Float).Value = drV("OPRICE")
                .Parameters.Add("ITEMAGE", SqlDbType.Int).Value = drV("ITEMAGE")
                .Parameters.Add("ITEMCOST", SqlDbType.Int).Value = drV("ITEMCOST")
                .Parameters.Add("ADMFLAG", SqlDbType.VarChar).Value = drV("ADMFLAG")
                .Parameters.Add("TAXFLAG", SqlDbType.VarChar).Value = drV("TAXFLAG")
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID 'Convert.ToString(sm.UserInfo.UserID)
                .ExecuteNonQuery()
            End With
        Next
    End Sub

    '3892
    Public Shared Sub SAVE_REVISE_ONCLASS(ByVal htSS As Hashtable, ByRef dtTemp As DataTable, ByVal oConn As SqlConnection)
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        Dim SCDate As String = TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        Dim SubSeqNo As String = TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO
        Dim ssUserID As String = TIMS.GetMyValue2(htSS, "ssUserID")
        If rPlanID = "" Then Exit Sub
        If rComIDNO = "" Then Exit Sub
        If rSeqNo = "" Then Exit Sub
        If SCDate = "" Then Exit Sub
        If SubSeqNo = "" Then Exit Sub
        If ssUserID = "" Then Exit Sub

        Call TIMS.OpenDbConn(oConn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO REVISE_ONCLASS (ROCID ,PLANID ,COMIDNO ,SEQNO ,SCDATE ,SUBSEQNO ,WEEKS ,TIMES ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        sql &= " VALUES (@ROCID ,@PLANID ,@COMIDNO ,@SEQNO ,@SCDATE ,@SUBSEQNO ,@WEEKS ,@TIMES ,@MODIFYACCT ,GETDATE())" & vbCrLf
        Dim iCmd As New SqlCommand(sql, oConn)

        sql = "" & vbCrLf
        sql &= " UPDATE REVISE_ONCLASS" & vbCrLf
        sql &= " SET WEEKS=@WEEKS ,TIMES=@TIMES ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE ROCID=@ROCID" & vbCrLf
        sql &= " AND PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SCDATE=@SCDATE AND SUBSEQNO=@SUBSEQNO" & vbCrLf
        Dim uCmd As New SqlCommand(sql, oConn)

        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM REVISE_ONCLASS WHERE ROCID=@ROCID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, oConn)

        sql = "" & vbCrLf
        sql &= " DELETE REVISE_ONCLASS WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SCDATE=@SCDATE AND SUBSEQNO=@SUBSEQNO" & vbCrLf
        Dim dCmd As New SqlCommand(sql, oConn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(rPlanID)
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = TIMS.CINT1(rSeqNo)
            .Parameters.Add("SCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SCDate)
            .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = TIMS.CINT1(SubSeqNo)
            .ExecuteNonQuery()
        End With

        For Each dr1 As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
            Dim FlagUseInsert1 As Boolean = False
            Dim FlagIsDelete As Boolean = False
            If dr1.RowState = DataRowState.Deleted Then FlagIsDelete = True

            Dim iROCID As Integer = 0 '= DbAccess.GetNewId(oConn, "REVISE_ONCLASS_ROCID_SEQ,REVISE_ONCLASS,ROCID")
            If Convert.ToString(dr1("ROCID")) = "" Then FlagUseInsert1 = True '空值-新增
            If Not FlagUseInsert1 Then iROCID = dr1("ROCID") '非新增 有值 修改
            If iROCID <= 0 Then
                '負數'取新值
                iROCID = DbAccess.GetNewId(oConn, "REVISE_ONCLASS_ROCID_SEQ,REVISE_ONCLASS,ROCID")
            End If
            If FlagIsDelete Then Continue For '刪除以下省略

            Dim dt As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("ROCID", SqlDbType.Int).Value = iROCID
                dt.Load(.ExecuteReader())
            End With
            If dt.Rows.Count <> 0 Then
                '修改
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("WEEKS", SqlDbType.VarChar).Value = dr1("WEEKS")
                    .Parameters.Add("TIMES", SqlDbType.VarChar).Value = dr1("TIMES")
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID

                    .Parameters.Add("ROCID", SqlDbType.Int).Value = iROCID
                    .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(rPlanID)
                    .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
                    .Parameters.Add("SEQNO", SqlDbType.Int).Value = TIMS.CINT1(rSeqNo)
                    .Parameters.Add("SCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SCDate)
                    .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = TIMS.CINT1(SubSeqNo)
                    .ExecuteNonQuery()
                End With
            Else
                '無資料的新增
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("ROCID", SqlDbType.Int).Value = iROCID
                    .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(rPlanID)
                    .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
                    .Parameters.Add("SEQNO", SqlDbType.Int).Value = TIMS.CINT1(rSeqNo)
                    .Parameters.Add("SCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SCDate)
                    .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = TIMS.CINT1(SubSeqNo)
                    .Parameters.Add("WEEKS", SqlDbType.VarChar).Value = dr1("WEEKS")
                    .Parameters.Add("TIMES", SqlDbType.VarChar).Value = dr1("TIMES")
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID
                    .ExecuteNonQuery()
                End With
            End If
        Next
    End Sub

    Public Shared Sub SAVE_REVISE_ONCLASS_OLD(ByVal htSS As Hashtable, ByRef dtTemp As DataTable, ByVal oConn As SqlConnection, ByRef DG1 As DataGrid)
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        Dim SCDate As String = TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        Dim SubSeqNo As String = TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO
        Dim ssUserID As String = TIMS.GetMyValue2(htSS, "ssUserID")
        If rPlanID = "" Then Exit Sub
        If rComIDNO = "" Then Exit Sub
        If rSeqNo = "" Then Exit Sub
        If SCDate = "" Then Exit Sub
        If SubSeqNo = "" Then Exit Sub
        If ssUserID = "" Then Exit Sub

        Call TIMS.OpenDbConn(oConn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'X'" & vbCrLf
        sql &= " FROM REVISE_ONCLASS_OLD" & vbCrLf
        sql &= " WHERE PLANID=@PLANID and COMIDNO=@COMIDNO and SEQNO=@SEQNO" & vbCrLf
        sql &= " and SCDATE=@SCDATE and SUBSEQNO=@SUBSEQNO" & vbCrLf
        sql &= " and WEEKS=@WEEKS and TIMES=@TIMES" & vbCrLf
        Dim sCmd As New SqlCommand(sql, oConn)
        sql = "" & vbCrLf
        sql &= " INSERT INTO REVISE_ONCLASS_OLD (ROCID ,PLANID ,COMIDNO ,SEQNO ,SCDATE ,SUBSEQNO ,WEEKS ,TIMES ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        sql &= " VALUES (@ROCID ,@PLANID ,@COMIDNO ,@SEQNO ,@SCDATE ,@SUBSEQNO ,@WEEKS ,@TIMES ,@MODIFYACCT ,GETDATE() )" & vbCrLf
        Dim iCmd As New SqlCommand(sql, oConn)
        For Each eItem As DataGridItem In DG1.Items
            Dim OldWeeks1 As Label = eItem.FindControl("OldWeeks1")
            Dim OldTimes1 As Label = eItem.FindControl("OldTimes1")
            OldWeeks1.Text = TIMS.ClearSQM(OldWeeks1.Text)
            OldTimes1.Text = TIMS.ClearSQM(OldTimes1.Text)
            Dim dt1 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(rPlanID)
                .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
                .Parameters.Add("SEQNO", SqlDbType.Int).Value = TIMS.CINT1(rSeqNo)
                .Parameters.Add("SCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SCDate)
                .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = TIMS.CINT1(SubSeqNo)

                .Parameters.Add("WEEKS", SqlDbType.NVarChar).Value = OldWeeks1.Text
                .Parameters.Add("TIMES", SqlDbType.NVarChar).Value = OldTimes1.Text
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count = 0 Then
                Dim iROCID As Integer = DbAccess.GetNewId(oConn, "REVISE_ONCLASS_OLD_ROCID_SEQ,REVISE_ONCLASS_OLD,ROCID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("ROCID", SqlDbType.Int).Value = iROCID
                    .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(rPlanID)
                    .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
                    .Parameters.Add("SEQNO", SqlDbType.Int).Value = TIMS.CINT1(rSeqNo)
                    .Parameters.Add("SCDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(SCDate)
                    .Parameters.Add("SUBSEQNO", SqlDbType.Int).Value = TIMS.CINT1(SubSeqNo)

                    .Parameters.Add("WEEKS", SqlDbType.NVarChar).Value = OldWeeks1.Text
                    .Parameters.Add("TIMES", SqlDbType.NVarChar).Value = OldTimes1.Text
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = ssUserID 'sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With
            End If
        Next
    End Sub

    Private Sub DataGrid21N_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid21New_1.ItemCommand, DataGrid21New_2.ItemCommand, DataGrid21New_3.ItemCommand, DataGrid21New_4.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim PCID As String = TIMS.GetMyValue(sCmdArg, "PCID")
        If PCID = "" Then Exit Sub
        If Session(Hid_COSTITEM_GUID21.Value) Is Nothing Then Exit Sub
        Dim dt As DataTable = Session(Hid_COSTITEM_GUID21.Value)
        If dt.Rows.Count = 0 Then Exit Sub
        Select Case e.CommandName
            Case cst_btnDel1Cmd
                ff33 = "PCID=" & PCID
                If dt.Select(ff33).Length > 0 Then dt.Select(ff33)(0).Delete()
                'dt.AcceptChanges()
                Session(Hid_COSTITEM_GUID21.Value) = dt
                Call SHOW_COSTITEM_GUID21(2)
        End Select
    End Sub

    Private Sub DataGrid21O_1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_1.ItemDataBound, DataGrid21New_1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim CostName As Label = e.Item.FindControl("CostName")
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim ItemCost As Label = e.Item.FindControl("ItemCost")
                Dim subtotal As Label = e.Item.FindControl("subtotal")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg
                CostName.Text = ""
                Select Case Convert.ToString(drv("CostID"))
                    Case "99"
                        strTMP1 = "其他-" & Convert.ToString(drv("ItemOther")).ToString
                    Case Else
                        ff33 = "CostID='" & drv("CostID") & "'"
                        If dt_KEY_COSTITEM.Select(ff33).Length > 0 Then strTMP1 = dt_KEY_COSTITEM.Select(ff33)(0)("CostName")
                End Select
                CostName.Text = strTMP1
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                ItemCost.Text = Convert.ToString(drv("ItemCost"))
                'subtotal.Text=TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
        End Select
    End Sub

    Private Sub DataGrid21O_2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_2.ItemDataBound, DataGrid21New_2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim ItemCost As Label = e.Item.FindControl("ItemCost")
                Dim subtotal As Label = e.Item.FindControl("subtotal")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                ItemCost.Text = Convert.ToString(drv("ItemCost"))
                'subtotal.Text=TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
        End Select
    End Sub

    Private Sub DataGrid21O_3_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_3.ItemDataBound, DataGrid21New_3.ItemDataBound
        'e.Item.Cells(4).Style("display")="none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim subtotal As Label = e.Item.FindControl("subtotal")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                'subtotal.Text=TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
        End Select
    End Sub

    Private Sub DataGrid21O_4_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_4.ItemDataBound, DataGrid21New_4.ItemDataBound
        'e.Item.Cells(5).Style("display")="none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim CostName As Label = e.Item.FindControl("CostName")
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim subtotal As Label = e.Item.FindControl("subtotal")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg
                CostName.Text = ""
                Select Case Convert.ToString(drv("CostID"))
                    Case "99"
                        strTMP1 = "其他-" & Convert.ToString(drv("ItemOther")).ToString
                    Case Else
                        ff33 = "CostID='" & drv("CostID") & "'"
                        If dt_KEY_COSTITEM.Select(ff33).Length > 0 Then strTMP1 = dt_KEY_COSTITEM.Select(ff33)(0)("CostName")
                End Select
                CostName.Text = strTMP1
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
        End Select
    End Sub

    '回上一頁
    Protected Sub btn_back_Click(sender As Object, e As EventArgs) Handles btn_back.Click
        Session(Hid_COSTITEM_GUID21.Value) = Nothing
        Session.Contents.Remove(Hid_COSTITEM_GUID21.Value)
        ViewState(vs_TEMP11_TrainDescDT) = ""
        ViewState(vs_TEMP20_TrainDescDT) = ""
        'Response.Redirect("TC_05_001.aspx?ID=" & Request("ID"))
        Dim url1 As String = "TC_05_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        Call TIMS.Utl_Redirect(Me, gobjconn, url1)
    End Sub

    ''' <summary>
    ''' (共用)將變更資料移動 至dr 欄位中
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="sm"></param>
    ''' <param name="iPlanKind"></param>
    ''' <param name="htSS"></param>
    Sub INS_CMN1(ByRef dr As DataRow, ByRef sm As SessionModel, ByRef iPlanKind As Integer, ByRef htSS As Hashtable)
        Dim v_ReviseCont As String = TIMS.ClearSQM(TIMS.GetMyValue2(htSS, "ReviseCont"))
        Dim v_changeReason As String = TIMS.ClearSQM(TIMS.GetMyValue2(htSS, "changeReason"))
        '變更-審核人員
        dr("ReviseAcct") = sm.UserInfo.UserID
        '變更內容說明
        dr("ReviseCont") = If(v_ReviseCont <> "", v_ReviseCont, Convert.DBNull)
        '變更原因說明
        dr("changeReason") = If(v_changeReason <> "", v_changeReason, Convert.DBNull)
        dr("Verifier") = If(iPlanKind = 1, sm.UserInfo.UserID, Convert.DBNull)
        dr("ReviseStatus") = If(iPlanKind = 1, "Y", Convert.DBNull)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()
    End Sub


    'https://jira.turbotech.com.tw/browse/TIMSC-249
    'Const cst_CostItemTable As String="CostItemTable"
    '計畫經費項目檔(PLAN_COSTITEM)
    'Public Shared Sub SHOW_COSTITEM_OLD(ByRef htSS As Hashtable, ByRef oConn As SqlConnection, ByRef dGrid2 As DataGrid)

    '計畫經費項目檔(PLAN_COSTITEM)
    Function GET_COSTITEMdt(ByRef htSS As Hashtable, ByRef oConn As SqlConnection) As DataTable
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") 'Request("PlanID")
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") 'Request("cid")
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") 'Request("no")
        'Dim iCostMode As Integer=TIMS.GetMyValue2(htSS, "CostMode") '計價方案'計價種類
        'Dim iPlanKind As Integer=TIMS.GetMyValue2(htSS, "PlanKind") '計價方案'計價種類
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM PLAN_COSTITEM"
        sql &= " WHERE PlanID='" & rPlanID & "'"
        sql &= " AND ComIDNO='" & rComIDNO & "'"
        sql &= " AND SeqNO='" & rSeqNo & "'"
        sql &= " ORDER BY PCID "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn)
        Return dt
    End Function

    '計畫經費項目檔(PLAN_COSTITEM)
    Sub SHOW_COSTITEM_1(ByRef htSS As Hashtable, ByRef oConn As SqlConnection)
        DataGrid21Old_1.Visible = False
        DataGrid21Old_2.Visible = False
        DataGrid21Old_3.Visible = False
        DataGrid21Old_4.Visible = False
        DataGrid21New_1.Visible = False
        DataGrid21New_2.Visible = False
        DataGrid21New_3.Visible = False
        DataGrid21New_4.Visible = False

        Dim objDG_Old As DataGrid = Nothing
        Dim objDG_New As DataGrid = Nothing
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") 'Request("PlanID")
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") 'Request("cid")
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") 'Request("no")
        Dim iPlanKind As Integer = TIMS.GetMyValue2(htSS, "iPlanKind") '
        Dim iCostMode As Integer = TIMS.GetMyValue2(htSS, "iCostMode") '計價方案'計價種類
        Dim dt1 As DataTable = GET_COSTITEMdt(htSS, oConn)
        Dim Total As Double = 0 '總費用
        Dim AdmTotal As Double = 0 '行政管理費 '(行政管理費百分比)
        Dim TaxTotal As Double = 0 '營業稅
        If dt1.Rows.Count = 0 Then Exit Sub
        'sql="SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'Dim iPlanKind As Integer=DbAccess.ExecuteScalar(sql, gobjconn)

        'PlanKind (1) 1:自辦 2:委外
        If iPlanKind = 1 Then
            objDG_Old = DataGrid21Old_1
            objDG_New = DataGrid21New_1
        Else
            Select Case iCostMode
                Case 2
                    '每人每時單價計價法
                    objDG_Old = DataGrid21Old_2
                    objDG_New = DataGrid21New_2
                Case 3
                    '每人輔助單價計價法
                    objDG_Old = DataGrid21Old_3
                    objDG_New = DataGrid21New_3
                Case 4
                    '個人單價計價法
                    objDG_Old = DataGrid21Old_4
                    objDG_New = DataGrid21New_4
            End Select
        End If

        objDG_Old.Visible = True
        With objDG_Old
            .DataSource = dt1
            .DataKeyField = "PCID"
            .DataBind()
        End With
        objDG_New.Visible = True
        With objDG_New
            .DataSource = dt1
            .DataKeyField = "PCID"
            .DataBind()
        End With
        Session(Hid_COSTITEM_GUID21.Value) = dt1
    End Sub

    '計畫經費項目檔(PLAN_COSTITEM) 計算金額-顯示
    Sub SHOW_COSTITEM_2(ByRef htSS As Hashtable, ByRef oConn As SqlConnection)
        'Dim rCDate As String=TIMS.GetMyValue2(htSS, "rCDate") 'Request("CDate")
        'Dim rSubNo As String=TIMS.GetMyValue2(htSS, "rSubNo") 'Request("SubNo")
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") 'Request("PlanID")
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") 'Request("cid")
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") 'Request("no")
        Dim iAdmPercent As Integer = TIMS.GetMyValue2(htSS, "iAdmPercent") '行政管理費 '(行政管理費百分比)
        Dim iTaxPercent As Integer = TIMS.GetMyValue2(htSS, "iTaxPercent") '營業稅 '(營業稅費用百分比)
        Dim iPlanKind As Integer = TIMS.GetMyValue2(htSS, "iPlanKind") '1:自辦  'PlanKind (2) 1:自辦 2:委外
        Dim iCostMode As Integer = TIMS.GetMyValue2(htSS, "iCostMode") '計價方案'計價種類

        Hid_PlanKind.Value = CStr(iPlanKind)
        Hid_CostMode.Value = CStr(iCostMode)
        Hid_AdmPercent.Value = CStr(iAdmPercent)
        Hid_TaxPercent.Value = CStr(iTaxPercent)

        'Dim ff33 As String=""
        Dim diTotal As Double = 0 '總費用(浮點數)
        Dim diAdmTotal As Double = 0 '行政管理費 '(行政管理費百分比)(浮點數)
        Dim diTaxTotal As Double = 0 '營業稅 '(營業稅費用百分比)(浮點數)
        Dim AdmCostText As String = "" '顯示文字
        Dim TaxCostText As String = "" '顯示文字
        'Dim flagAdmGrantTR As Boolean=False 'false:預設不顯示 True '.Visible=True
        'Dim flagTaxGrantTR As Boolean=False 'false:預設不顯示'營業稅 '(營業稅費用百分比)
        AdmGrantTROld.Visible = False
        TaxGrantTROld.Visible = False
        AdmGrantTRNew.Visible = False
        TaxGrantTRNew.Visible = False

        'Dim ff As String=""
        'Dim sql As String=""
        'sql=""
        'sql &= " SELECT * FROM REVISE_COSTITEM_OLD"
        'sql &= " WHERE 1=1"
        'sql &= " and PlanID=@PlanID"
        'sql &= " and ComIDNO=@ComIDNO"
        'sql &= " and SeqNO=@SeqNO"
        'sql &= " and CDate=dbo.fn_DATE(@CDate)" & vbCrLf
        'sql &= " and SubSeqNo=@SubSeqNo" & vbCrLf
        'sql &= " ORDER BY COSTID"
        'Dim sCmd As New SqlCommand(sql, oConn)
        'Dim dt1 As New DataTable
        'With sCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("PlanID", SqlDbType.VarChar).Value=rPlanID
        '    .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value=rComIDNO
        '    .Parameters.Add("SeqNO", SqlDbType.VarChar).Value=rSeqNo
        '    .Parameters.Add("CDate", SqlDbType.VarChar).Value=rCDate
        '    .Parameters.Add("SubSeqNo", SqlDbType.VarChar).Value=rSubNo
        '    dt1.Load(.ExecuteReader())
        'End With
        Dim dt1 As DataTable = GET_COSTITEMdt(htSS, oConn)
        Session(Hid_COSTITEM_GUID21.Value) = dt1
        Call SHOW_COSTITEM_GUID21(1)
        Session(Hid_COSTITEM_GUID21.Value) = dt1
        Call SHOW_COSTITEM_GUID21(2)
    End Sub

    '取得申請變更的資訊
    Sub SHOW_REVISE_COSTITEM(ByVal dt3 As DataTable, ByVal iRCID As Integer, ByVal iO1N2 As Integer)
        Dim iPlanKind As Integer = TIMS.CINT1(Hid_PlanKind.Value) '
        Dim iCostMode As Integer = TIMS.CINT1(Hid_CostMode.Value) '計價方案'計價種類
        Select Case iO1N2
            Case 1
                Dim objDG_Old As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_Old = DataGrid21Old_1
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_Old = DataGrid21Old_2
                        Case 3
                            '每人輔助單價計價法
                            objDG_Old = DataGrid21Old_3
                        Case 4
                            '個人單價計價法
                            objDG_Old = DataGrid21Old_4
                    End Select
                End If
                objDG_Old.Visible = True
                With objDG_Old
                    .DataSource = dt3
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                Session(Hid_COSTITEM_GUID21.Value) = dt3
                Call SHOW_COSTITEM_GUID21(1)
            Case 2
                Dim objDG_New As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_New = DataGrid21New_1
                    objDG_New.Columns(5).Visible = False
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_New = DataGrid21New_2
                            objDG_New.Columns(4).Visible = False
                        Case 3
                            '每人輔助單價計價法
                            objDG_New = DataGrid21New_3
                            objDG_New.Columns(3).Visible = False
                        Case 4
                            '個人單價計價法
                            objDG_New = DataGrid21New_4
                            objDG_New.Columns(4).Visible = False
                    End Select
                End If
                objDG_New.Visible = True
                With objDG_New
                    .DataSource = dt3
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                Session(Hid_COSTITEM_GUID21.Value) = dt3
                Call SHOW_COSTITEM_GUID21(2)
        End Select
    End Sub

    '計畫經費項目檔(PLAN_COSTITEM) 依SESSION 
    Sub SHOW_COSTITEM_GUID21(ByVal iO1N2 As Integer)
        If Session(Hid_COSTITEM_GUID21.Value) Is Nothing Then Exit Sub
        Dim dt1 As DataTable = Session(Hid_COSTITEM_GUID21.Value)
        Dim iPlanKind As Integer = TIMS.CINT1(Hid_PlanKind.Value)
        Dim iCostMode As Integer = TIMS.CINT1(Hid_CostMode.Value) '計價方案'計價種類
        Dim iAdmPercent As Integer = TIMS.CINT1(Hid_AdmPercent.Value) '行政管理費 '(行政管理費百分比)
        Dim iTaxPercent As Integer = TIMS.CINT1(Hid_TaxPercent.Value) '營業稅 '(營業稅費用百分比)

        Dim ff As String = ""
        Dim diTotal As Double = 0 '總費用(浮點數)
        Dim diAdmTotal As Double = 0 '行政管理費 '(行政管理費百分比)(浮點數)
        Dim diTaxTotal As Double = 0 '營業稅 '(營業稅費用百分比)(浮點數)
        Dim AdmCostText As String = "" '顯示文字
        Dim TaxCostText As String = "" '顯示文字
        Const cst_Plankind1t As String = "費用列表"
        Const cst_CostMode2t As String = "每人每時計價"
        Const cst_CostMode3t As String = "每人輔助計價"
        Const cst_CostMode4t As String = "個人單價計價"

        Select Case iO1N2
            Case 1
                'PlanKind (2) 1:自辦 2:委外
                If iPlanKind = 1 Then
                    labcost21txt1Old.Text = cst_Plankind1t
                Else
                    Select Case iCostMode
                        Case 2 '每人每時單價計價法-2:委外
                            labcost21txt1Old.Text = cst_CostMode2t
                        Case 3 '每人輔助單價計價法-2:委外
                            labcost21txt1Old.Text = cst_CostMode3t
                        Case 4 '個人單價計價法-2:委外
                            labcost21txt1Old.Text = cst_CostMode4t
                    End Select
                End If
                DataGrid21Old_1.Visible = False
                DataGrid21Old_2.Visible = False
                DataGrid21Old_3.Visible = False
                DataGrid21Old_4.Visible = False
                Dim objDG_Old As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_Old = DataGrid21Old_1
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_Old = DataGrid21Old_2
                        Case 3
                            '每人輔助單價計價法
                            objDG_Old = DataGrid21Old_3
                        Case 4
                            '個人單價計價法
                            objDG_Old = DataGrid21Old_4
                    End Select
                End If
                objDG_Old.Visible = True
                With objDG_Old
                    .DataSource = dt1
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                Session(Hid_COSTITEM_GUID21.Value) = dt1
            Case 2
                'PlanKind (2) 1:自辦 2:委外
                If iPlanKind = 1 Then
                    labcost21txt1New.Text = cst_Plankind1t
                Else
                    Select Case iCostMode
                        Case 2 '每人每時單價計價法-2:委外
                            labcost21txt1New.Text = cst_CostMode2t
                        Case 3 '每人輔助單價計價法-2:委外
                            labcost21txt1New.Text = cst_CostMode3t
                        Case 4 '個人單價計價法-2:委外
                            labcost21txt1New.Text = cst_CostMode4t
                    End Select
                End If
                DataGrid21New_1.Visible = False
                DataGrid21New_2.Visible = False
                DataGrid21New_3.Visible = False
                DataGrid21New_4.Visible = False
                Dim objDG_New As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_New = DataGrid21New_1
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_New = DataGrid21New_2
                        Case 3
                            '每人輔助單價計價法
                            objDG_New = DataGrid21New_3
                        Case 4
                            '個人單價計價法
                            objDG_New = DataGrid21New_4
                    End Select
                End If
                objDG_New.Visible = True
                With objDG_New
                    .DataSource = dt1
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                Session(Hid_COSTITEM_GUID21.Value) = dt1
        End Select

        If iPlanKind = 1 Then
            '1:自辦  'PlanKind (2) 1:自辦 2:委外
            '行政管理費 '(行政管理費百分比)
            If iAdmPercent > -1 Then
                Dim aFlag As Boolean = False 'false:未啟用行政管理費百分比
                ff = "AdmFlag='Y'"
                If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用行政管理費百分比
                If aFlag Then
                    Dim strTMP1 As String = ""
                    Dim fff As String = "AdmFlag='Y'"
                    For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                        diAdmTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                        If strTMP1 <> "" Then strTMP1 &= "+"
                        Select Case Convert.ToString(drv("CostID"))
                            Case "99"
                                strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                            Case Else
                                ff = "CostID='" & drv("CostID") & "'"
                                If dt_KEY_COSTITEM.Select(ff).Length > 0 Then strTMP1 &= dt_KEY_COSTITEM.Select(ff)(0)("CostName")
                        End Select
                    Next
                    AdmCostText = "(" & strTMP1 & ")*" & iAdmPercent & "%=" & TIMS.ROUND(diAdmTotal * iAdmPercent / 100)
                    Select Case iO1N2
                        Case 1
                            AdmGrantTROld.Visible = True 'true:顯示 
                            AdmCostOld.Text = AdmCostText
                        Case 2
                            AdmGrantTRNew.Visible = True 'true:顯示 
                            AdmCostNew.Text = AdmCostText
                    End Select
                End If
            End If

            '營業稅 '(營業稅費用百分比)
            If iTaxPercent > -1 Then
                Dim aFlag As Boolean = False 'false:未啟用 營業稅費用百分比
                ff = "TaxFlag='Y'"
                If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用 營業稅費用百分比
                If aFlag Then
                    Dim strTMP1 As String = ""
                    Dim fff As String = "TaxFlag='Y'"
                    For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                        diTaxTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                        If strTMP1 <> "" Then strTMP1 &= "+"
                        Select Case Convert.ToString(drv("CostID"))
                            Case "99"
                                strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                            Case Else
                                ff = "CostID='" & drv("CostID") & "'"
                                If dt_KEY_COSTITEM.Select(ff).Length > 0 Then strTMP1 &= dt_KEY_COSTITEM.Select(ff)(0)("CostName")
                        End Select
                    Next
                    TaxCostText = "(" & strTMP1 & ")*" & iTaxPercent & "%=" & TIMS.ROUND(diTaxTotal * iTaxPercent / 100)
                    Select Case iO1N2
                        Case 1
                            TaxGrantTROld.Visible = True 'true:顯示 
                            TaxCostOld.Text = TaxCostText
                        Case 2
                            TaxGrantTRNew.Visible = True 'true:顯示 
                            TaxCostNew.Text = TaxCostText
                    End Select
                End If
            End If

            diTotal = 0
            For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
            Next
            Select Case iO1N2
                Case 1
                    '行政管理費 '(行政管理費百分比)
                    If AdmGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                    '營業稅 '(營業稅費用百分比)
                    If TaxGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                    TotalCost1Old.Text = TIMS.ROUND(diTotal)
                Case 2
                    '行政管理費 '(行政管理費百分比)
                    If AdmGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                    '營業稅 '(營業稅費用百分比)
                    If TaxGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                    TotalCost1New.Text = TIMS.ROUND(diTotal)
            End Select
        Else
            'PlanKind (2) 1:自辦 2:委外
            Select Case iCostMode
                Case 2 '每人每時單價計價法-2:委外
                    diTotal = 0
                    For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                    Next
                    diTotal = CDbl(TIMS.ROUND(diTotal))
                Case 3 '每人輔助單價計價法-2:委外
                    diTotal = 0
                    For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
                    Next
                    diTotal = CDbl(TIMS.ROUND(diTotal))
                Case 4 '個人單價計價法-2:委外
                    '行政管理費 '(行政管理費百分比)
                    If iAdmPercent > -1 Then
                        Dim aFlag As Boolean = False 'false:未啟用行政管理費百分比
                        ff = "AdmFlag='Y'"
                        If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用行政管理費百分比
                        If aFlag Then
                            Dim strTMP1 As String = ""
                            Dim fff As String = "AdmFlag='Y'"
                            For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                                diAdmTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                                If strTMP1 <> "" Then strTMP1 &= "+"
                                Select Case Convert.ToString(drv("CostID"))
                                    Case "99"
                                        strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                                    Case Else
                                        ff = "CostID='" & drv("CostID") & "'"
                                        If dt_KEY_COSTITEM.Select(ff).Length > 0 Then strTMP1 &= dt_KEY_COSTITEM.Select(ff)(0)("CostName")
                                End Select
                            Next
                            AdmCostText = "(" & strTMP1 & ")*" & iAdmPercent & "%=" & TIMS.ROUND(diAdmTotal * iAdmPercent / 100)
                            Select Case iO1N2
                                Case 1
                                    AdmGrantTROld.Visible = True 'true:顯示 
                                    AdmCostOld.Text = TaxCostText
                                Case 2
                                    AdmGrantTRNew.Visible = True 'true:顯示 
                                    AdmCostNew.Text = AdmCostText
                            End Select
                        End If
                    End If
                    '營業稅 '(營業稅費用百分比)
                    If iTaxPercent > -1 Then
                        Dim aFlag As Boolean = False 'false:未啟用 營業稅費用百分比
                        ff = "TaxFlag='Y'"
                        If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用 營業稅費用百分比
                        If aFlag Then
                            Dim strTMP1 As String = ""
                            Dim fff As String = "TaxFlag='Y'"
                            For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                                diTaxTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                                If strTMP1 <> "" Then strTMP1 &= "+"
                                Select Case Convert.ToString(drv("CostID"))
                                    Case "99"
                                        strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                                    Case Else
                                        ff = "CostID='" & drv("CostID") & "'"
                                        If dt_KEY_COSTITEM.Select(ff).Length > 0 Then strTMP1 &= dt_KEY_COSTITEM.Select(ff)(0)("CostName")
                                End Select
                            Next
                            TaxCostText = "(" & strTMP1 & ")*" & iTaxPercent & "%=" & TIMS.ROUND(diTaxTotal * iTaxPercent / 100)
                            Select Case iO1N2
                                Case 1
                                    TaxGrantTROld.Visible = True 'true:顯示 
                                    TaxCostOld.Text = TaxCostText
                                Case 2
                                    TaxGrantTRNew.Visible = True 'true:顯示 
                                    TaxCostNew.Text = TaxCostText
                            End Select
                        End If
                    End If
                    diTotal = 0
                    For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
                    Next
                    Select Case iO1N2
                        Case 1
                            '行政管理費 '(行政管理費百分比)
                            If AdmGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                            '營業稅 '(營業稅費用百分比)
                            If TaxGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                            TotalCost1Old.Text = TIMS.ROUND(diTotal)
                        Case 2
                            '行政管理費 '(行政管理費百分比)
                            If AdmGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                            '營業稅 '(營業稅費用百分比)
                            If TaxGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                            TotalCost1New.Text = TIMS.ROUND(diTotal)
                    End Select
            End Select
        End If
    End Sub

    '產生新的GUID 避免記憶體相同 而異常
    Sub CREATE_NEW_GUID21()
        'If IsPostBack Then Exit Sub
        Hid_COSTITEM_GUID21.Value = TIMS.GetGUID()
        If Session(cst_TC05001CHG_COSTITEM_GUID) IsNot Nothing Then
            '清理上一個SESSION GUID
            Dim TC05001CHG_COSTITEM_GUID As String = Session(cst_TC05001CHG_COSTITEM_GUID)
            Session(TC05001CHG_COSTITEM_GUID) = Nothing
            Session.Contents.Remove(TC05001CHG_COSTITEM_GUID)
        End If
        '記錄這一個SESSION GUID
        Session(cst_TC05001CHG_COSTITEM_GUID) = Hid_COSTITEM_GUID21.Value
    End Sub

    Private Sub DataGrid21_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid21.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim Specialty1 As Label = e.Item.FindControl("Specialty1")
                'Dim ProLicense As Label=e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEA As HtmlInputButton = e.Item.FindControl("btn_TCTYPEA")
                Dim rqRID As String = sm.UserInfo.RID
                If RIDValue.Value.Length > 1 Then rqRID = RIDValue.Value
                sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                btn_TCTYPEA.Attributes("onclick") = sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                Specialty1.Text = Convert.ToString(drv("Specialty1"))
                'ProLicense.Text=Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'If Hid_NewData11_3.Value <> "" Then TeacherDesc.Text=Hid_NewData11_3.Value 'Convert.ToString(drv("TeacherDesc"))

                TeacherDesc.ReadOnly = False
                btn_TCTYPEA.Visible = True

                Dim flag_can_save As Boolean = True
                If RIDValue.Value = "" Then flag_can_save = False '不同單位 不提供儲存
                If sm.UserInfo.RID <> RIDValue.Value Then flag_can_save = False '不同單位 不提供儲存
                Select Case sm.UserInfo.LID
                    Case 2
                        '不同單位 不提供儲存
                        If Not flag_can_save Then
                            TeacherDesc.ReadOnly = True
                            btn_TCTYPEA.Visible = False
                        End If
                End Select
                Select Case rActCheck
                    Case Cst_cRevise
                        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
                        If Hid_PARTREDUC_Y_CanUpdate.Value = "" Then
                            TeacherDesc.ReadOnly = True
                            btn_TCTYPEA.Visible = False
                        End If
                End Select

                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly=True
                '        btn_TCTYPEA.Visible=False
                'End Select

        End Select
    End Sub

    Private Sub DataGrid22_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid22.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim Specialty1 As Label = e.Item.FindControl("Specialty1")
                'Dim ProLicense As Label=e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEB As HtmlInputButton = e.Item.FindControl("btn_TCTYPEB")
                Dim rqRID As String = sm.UserInfo.RID '使用者單位
                If RIDValue.Value.Length > 1 Then rqRID = RIDValue.Value
                sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                btn_TCTYPEB.Attributes("onclick") = sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                Specialty1.Text = Convert.ToString(drv("Specialty1"))
                'ProLicense.Text=Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'If Hid_NewData20_3.Value <> "" Then TeacherDesc.Text=Hid_NewData20_3.Value 'Convert.ToString(drv("TeacherDesc"))

                TeacherDesc.ReadOnly = False
                btn_TCTYPEB.Visible = True

                Dim flag_can_save As Boolean = True
                If RIDValue.Value = "" Then flag_can_save = False '不同單位 不提供儲存
                If sm.UserInfo.RID <> RIDValue.Value Then flag_can_save = False '不同單位 不提供儲存
                Select Case sm.UserInfo.LID
                    Case 2
                        '不同單位 不提供儲存
                        If Not flag_can_save Then
                            TeacherDesc.ReadOnly = True
                            btn_TCTYPEB.Visible = False
                        End If
                End Select
                Select Case rActCheck
                    Case Cst_cRevise
                        'OJT-21080201：產投 -班級變更審核：新增「還原」按鈕
                        If Hid_PARTREDUC_Y_CanUpdate.Value = "" Then
                            TeacherDesc.ReadOnly = True
                            btn_TCTYPEB.Visible = False
                        End If
                End Select

                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly=True
                '        btn_TCTYPEB.Visible=False
                'End Select
        End Select
    End Sub


    '建立可選教師列表-遴選辦法說明
    Sub SHOW_REVISE_TEACHER12(ByRef htSS As Hashtable, ByRef oConn As SqlConnection)
        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        Dim SCDate As String = TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        Dim SubSeqNo As String = TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO
        Dim vActCheck As String = TIMS.GetMyValue2(htSS, "ActCheck") 'RActCheck / Cst_cPlan '申請 /Cst_cRevise '審核查詢

        Dim rqRID As String = TIMS.GetMyValue2(htSS, "RID")
        Dim rqTECHIDs As String = TIMS.GetMyValue2(htSS, "TECHIDs")
        Dim TechTYPE As String = TIMS.GetMyValue2(htSS, "TechTYPE") 'A/B
        If rqTECHIDs = "" Then Exit Sub
        Dim inTECHIDs As String = TIMS.CombiSQM2IN(rqTECHIDs)
        If inTECHIDs = "" Then Exit Sub

        Dim parms As New Hashtable
        Dim sql As String = ""

        Select Case vActCheck 'UCase(Request("check"))
            Case Cst_cPlan '申請
                sql = "" & vbCrLf
                sql &= " SELECT a.TechID" & vbCrLf '教師ID
                sql &= " ,a.TeachCName" & vbCrLf '教師姓名 
                sql &= " ,a.DegreeID" & vbCrLf '學歷
                sql &= " ,c.Name DegreeName" & vbCrLf '學歷
                '專業領域 Specialty1
                sql &= " ,ISNULL(a.Specialty1, '') Specialty1" & vbCrLf
                '專業證照-相關證照
                sql &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2" & vbCrLf
                sql &= " ELSE a.ProLicense END ProLicense" & vbCrLf
                '遴選辦法說明
                sql &= " ,convert(nvarchar(500),null) TeacherDesc " 'TechTYPE: A:師資/B:助教
                'sql &= " ,dbo.fn_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, '" & TechTYPE & "', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
                sql &= " FROM TEACH_TEACHERINFO a" & vbCrLf
                sql &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID=c.DegreeID" & vbCrLf
                sql &= " WHERE a.WorkStatus='1'" & vbCrLf
                sql &= " AND a.RID='" & rqRID & "'" & vbCrLf
                sql &= " AND a.TechID IN (" & inTECHIDs & ")" & vbCrLf
                sql &= " ORDER BY a.TechID" & vbCrLf

                parms.Clear()

            Case Cst_cRevise '審核查詢
                'Dim sql As String=""
                sql = ""
                sql &= " WITH WP1 AS (" & vbCrLf
                sql &= "  SELECT PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO" & vbCrLf
                sql &= "  FROM PLAN_REVISE" & vbCrLf
                sql &= "  WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                sql &= "  AND CDATE =@CDATE AND SUBSEQNO=@SUBSEQNO" & vbCrLf
                'AND PLANID='4818' AND COMIDNO='80592907' AND SEQNO='1'	AND CDATE ='2020-04-07'	AND SUBSEQNO=1
                sql &= " )" & vbCrLf
                sql &= " SELECT a.TechID" & vbCrLf '教師ID
                sql &= " ,a.TeachCName" & vbCrLf '教師姓名 
                sql &= " ,a.DegreeID" & vbCrLf '學歷
                sql &= " ,c.Name DegreeName" & vbCrLf '學歷
                '專業領域 Specialty1
                sql &= " ,ISNULL(a.Specialty1, '') Specialty1" & vbCrLf
                '專業證照-相關證照
                sql &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1+'、'+a.ProLicense2" & vbCrLf
                sql &= " ELSE a.ProLicense END ProLicense" & vbCrLf
                '遴選辦法說明
                sql &= " ,dbo.FN_GET_REVISE_TEACHER3(b.PLANID, b.COMIDNO, b.SEQNO,b.CDATE, b.SUBSEQNO, '" & TechTYPE & "', a.TechID) TeacherDesc" & vbCrLf 'TechTYPE: A:師資/B:助教
                sql &= " FROM TEACH_TEACHERINFO a" & vbCrLf
                sql &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID=c.DegreeID" & vbCrLf
                'CROSS JOIN
                sql &= " CROSS JOIN WP1 b" & vbCrLf
                sql &= " WHERE a.WorkStatus='1' AND a.RID='" & rqRID & "'" & vbCrLf
                sql &= " AND a.TechID IN (" & inTECHIDs & ")" & vbCrLf
                'sql &= " AND a.RID='F3962'" & vbCrLf 'sql &= " AND a.TechID IN (432456,441777,441706,441771,435528,441840,441696)" & vbCrLf
                sql &= " ORDER BY a.TechID" & vbCrLf

                parms.Clear()
                parms.Add("PLANID", TIMS.CINT1(rPlanID))
                parms.Add("COMIDNO", rComIDNO)
                parms.Add("SEQNO", TIMS.CINT1(rSeqNo))
                parms.Add("CDATE", TIMS.Cdate2(SCDate))
                parms.Add("SUBSEQNO", TIMS.CINT1(SubSeqNo))
        End Select

        Select Case TechTYPE
            Case "A" '師資
                Dim dtT As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
                i_gSeqno = 0
                tbDataGrid21.Visible = False
                If dtT.Rows.Count > 0 Then
                    tbDataGrid21.Visible = True
                    DataGrid21.DataSource = dtT
                    DataGrid21.DataBind()
                End If

            Case "B" '助教
                Dim dtT2 As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
                i_gSeqno = 0
                tbDataGrid22.Visible = False
                If dtT2.Rows.Count > 0 Then
                    tbDataGrid22.Visible = True
                    DataGrid22.DataSource = dtT2
                    DataGrid22.DataBind()
                End If
        End Select

    End Sub

    '檢查 班級申請老師
    Function CHK_REVISE_TEACHER(ByRef htSS As Hashtable, ByRef errmsg As String) As Boolean
        '#Region "檢查 班級申請老師"
        Dim rst As Boolean = True
        Const Cst_授課教師限制數 As Integer = 0 '10 '0:無限制

        'Dim rPlanID As String=TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        'Dim rComIDNO As String=TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        'Dim rSeqNo As String=TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        'Dim SCDate As String=TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        'Dim SubSeqNo As String=TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO
        'Dim vActCheck As String=TIMS.GetMyValue2(htSS, "ActCheck") 'RActCheck / Cst_cPlan '申請 /Cst_cRevise '審核查詢
        'Dim rqRID As String=TIMS.GetMyValue2(htSS, "RID")
        Dim rqTECHIDs As String = TIMS.GetMyValue2(htSS, "TECHIDs")
        Dim TechTYPE As String = TIMS.GetMyValue2(htSS, "TechTYPE") 'A/B
        If rqTECHIDs = "" Then errmsg &= "請選擇有效資料" & vbCrLf
        Dim inTECHIDs As String = TIMS.CombiSQM2IN(rqTECHIDs)
        If inTECHIDs = "" Then errmsg &= "請選擇有效資料" & vbCrLf
        If errmsg <> "" Then Return False

        Select Case TechTYPE
            Case "A"
                Dim i As Integer = 0
                Dim errT As String = ""
                Dim i_errI2 As Integer = 0
                For Each eItem As DataGridItem In DataGrid21.Items
                    'Dim HidTechID As HtmlInputHidden=eItem.FindControl("HidTechID")
                    Dim seqno As Label = eItem.FindControl("seqno")
                    Dim TeachCName As Label = eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    'Dim btn_TCTYPEA As HtmlInputButton=eItem.FindControl("btn_TCTYPEA") 'TechTYPE: A:師資/B:助教
                    i += 1
                    TeacherDesc.Text = TIMS.ClearSQM(TeacherDesc.Text)
                    If TeacherDesc.Text = "" Then
                        errT = seqno.Text & ":" & TeachCName.Text
                        i_errI2 += 1
                        Exit For
                    End If
                Next
                If i = 0 Then
                    errmsg &= "至少選擇1筆授課教師" & vbCrLf
                    Return False
                End If
                If i_errI2 > 0 Then
                    errmsg &= "授課教師-" & errT & "-遴選辦法說明辦法為必填" & vbCrLf
                    Return False
                End If
                If Cst_授課教師限制數 <> 0 Then '0:無限制
                    If Not (i <= Cst_授課教師限制數) Then
                        errmsg &= "僅可選擇" & Cst_授課教師限制數 & "筆授課教師" & vbCrLf
                        Return False
                    End If
                End If

            Case "B"
                Dim errTB As String = ""
                Dim i_errI2B As Integer = 0
                For Each eItem As DataGridItem In DataGrid22.Items
                    'Dim HidTechID As HtmlInputHidden=eItem.FindControl("HidTechID")
                    Dim seqno As Label = eItem.FindControl("seqno")
                    Dim TeachCName As Label = eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    'Dim btn_TCTYPEB As HtmlInputButton=eItem.FindControl("btn_TCTYPEB") 'TechTYPE: A:師資/B:助教
                    TeacherDesc.Text = TIMS.ClearSQM(TeacherDesc.Text)
                    If TeacherDesc.Text = "" Then
                        errTB = seqno.Text & ":" & TeachCName.Text
                        i_errI2B += 1
                        Exit For
                    End If
                Next
                If i_errI2B > 0 Then
                    errmsg &= "授課助教-" & errTB & "-遴選辦法說明辦法為必填" & vbCrLf
                    Return False
                End If
        End Select

        Return rst
    End Function

    '儲存 班級申請老師-PLAN_TEACHER
    Sub SAVE_REVISE_TEACHER(ByRef htSS As Hashtable, ByVal tConn As SqlConnection)
        'Dim rst As String=""

        Dim rPlanID As String = TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        Dim rComIDNO As String = TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        Dim rSeqNo As String = TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        Dim SCDate As String = TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        Dim SubSeqNo As String = TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO
        'Dim rqRID As String=TIMS.GetMyValue2(htSS, "RID")
        Dim rqTECHIDs As String = TIMS.GetMyValue2(htSS, "TECHIDs")
        Dim TechTYPE As String = TIMS.GetMyValue2(htSS, "TechTYPE") 'A/B
        If rqTECHIDs = "" Then Exit Sub
        'Dim inTECHIDs As String=TIMS.CombiSQM2IN(rqTECHIDs)
        'If inTECHIDs="" Then Exit Sub

        Dim iSql As String = ""
        iSql = ""
        iSql &= " INSERT INTO REVISE_TEACHER ( PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,TECHID,TECHTYPE ,TEACHERDESC,MODIFYACCT,MODIFYDATE)" & vbCrLf
        iSql &= " VALUES ( @PLANID,@COMIDNO,@SEQNO,@CDATE,@SUBSEQNO,@TECHID,@TECHTYPE ,@TEACHERDESC,@MODIFYACCT,GETDATE())" & vbCrLf

        Dim i_have_data As Integer = 0
        Select Case TechTYPE
            Case "A"
                For Each eItem As DataGridItem In DataGrid21.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    TeacherDesc.Text = TIMS.ClearSQM(TeacherDesc.Text)
                    If HidTechID.Value <> "" AndAlso TeacherDesc.Text <> "" Then
                        i_have_data += 1
                        Exit For
                    End If
                Next
            Case "B"
                For Each eItem As DataGridItem In DataGrid22.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    TeacherDesc.Text = TIMS.ClearSQM(TeacherDesc.Text)
                    If HidTechID.Value <> "" AndAlso TeacherDesc.Text <> "" Then
                        i_have_data += 1
                        Exit For
                    End If
                Next
        End Select

        If i_have_data > 0 Then
            '有資料才可以刪除
            Dim dSql As String = ""
            dSql &= " DELETE REVISE_TEACHER "
            dSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
            dSql &= " AND CDATE =@CDATE AND SUBSEQNO=@SUBSEQNO" & vbCrLf
            dSql &= " AND TechTYPE=@TechTYPE" & vbCrLf
            Dim d_parms As New Hashtable
            With d_parms
                .Clear()
                .Add("PLANID", TIMS.CINT1(rPlanID))
                .Add("COMIDNO", rComIDNO)
                .Add("SEQNO", TIMS.CINT1(rSeqNo))
                .Add("CDATE", TIMS.Cdate2(SCDate))
                .Add("SUBSEQNO", TIMS.CINT1(SubSeqNo))
                .Add("TechTYPE", TechTYPE)
            End With
            DbAccess.ExecuteNonQuery(dSql, tConn, d_parms)
        End If

        'Dim tTEACHERDESC As String=""
        'tTEACHERDESC=TeacherDesc_A.Text
        'Dim tTECHTYPE As String="" 'TechTYPE: A:師資/B:助教
        'tTECHTYPE="A" 'TechTYPE: A:師資/B:助教
        Select Case TechTYPE
            Case "A"
                For Each eItem As DataGridItem In DataGrid21.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    'Dim seqno As Label=eItem.FindControl("seqno")
                    'Dim TeachCName As Label=eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    'Dim btn_TCTYPEA As HtmlInputButton=eItem.FindControl("btn_TCTYPEA")
                    Dim tTEACHERDESC As String = TIMS.ClearSQM(TeacherDesc.Text)
                    If tTEACHERDESC.Length > 500 Then tTEACHERDESC = tTEACHERDESC.Substring(0, 500)
                    If HidTechID.Value <> "" Then
                        Dim parms As New Hashtable
                        'parms.Clear()
                        parms.Add("PLANID", TIMS.CINT1(rPlanID))
                        parms.Add("COMIDNO", rComIDNO)
                        parms.Add("SEQNO", TIMS.CINT1(rSeqNo))
                        parms.Add("CDATE", TIMS.Cdate2(SCDate))
                        parms.Add("SUBSEQNO", TIMS.CINT1(SubSeqNo))
                        parms.Add("TECHID", TIMS.CINT1(HidTechID.Value)) 'dr("TECHID"))
                        parms.Add("TECHTYPE", TechTYPE) 'tTECHTYPE) 'TechTYPE: A:師資/B:助教

                        parms.Add("TEACHERDESC", tTEACHERDESC)
                        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                        DbAccess.ExecuteNonQuery(iSql, tConn, parms)
                    End If
                Next
                'Return rst

            Case "B"
                'tTECHTYPE="B" 'TechTYPE: A:師資/B:助教
                For Each eItem As DataGridItem In DataGrid22.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    'Dim seqno As Label=eItem.FindControl("seqno")
                    'Dim TeachCName As Label=eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    'Dim btn_TCTYPEB As HtmlInputButton=eItem.FindControl("btn_TCTYPEB")
                    Dim tTEACHERDESC As String = TIMS.ClearSQM(TeacherDesc.Text)
                    If tTEACHERDESC.Length > 500 Then tTEACHERDESC = tTEACHERDESC.Substring(0, 500)
                    If HidTechID.Value <> "" Then
                        Dim parms As New Hashtable
                        'parms.Clear()
                        parms.Add("PLANID", TIMS.CINT1(rPlanID))
                        parms.Add("COMIDNO", rComIDNO)
                        parms.Add("SEQNO", TIMS.CINT1(rSeqNo))
                        parms.Add("CDATE", TIMS.Cdate2(SCDate))
                        parms.Add("SUBSEQNO", TIMS.CINT1(SubSeqNo))
                        parms.Add("TECHID", TIMS.CINT1(HidTechID.Value)) 'dr("TECHID"))
                        parms.Add("TECHTYPE", TechTYPE) 'tTECHTYPE) 'TechTYPE: A:師資/B:助教

                        parms.Add("TEACHERDESC", tTEACHERDESC)
                        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                        DbAccess.ExecuteNonQuery(iSql, tConn, parms)
                    End If
                Next
                'Return rst
        End Select
        'Return "沒有儲存活動？"
    End Sub

    'GET_OldDataVal OldData14_1b.Value ==dr("OldData14_1")/dr("NewData14_1")/NewData14_1b/SciPlaceIDb 
    ''' <summary>
    ''' 顯示文字 顯示下拉
    ''' </summary>
    ''' <param name="v_OldData"></param>
    ''' <param name="v_NewData"></param>
    ''' <param name="o_NewDDL1"></param>
    ''' <param name="o_Label1"></param>
    ''' <returns></returns>
    Public Shared Function GET_OldDataVal(ByVal v_OldData As String, ByVal v_NewData As String, ByRef o_NewDDL1 As DropDownList, ByRef o_Label1 As Label, ByVal flag_Enabled As Boolean) As String
        o_Label1.Text = If(o_NewDDL1.Items.FindByValue(v_OldData) IsNot Nothing, o_NewDDL1.Items.FindByValue(v_OldData).Text, "") '舊值顯示文字
        '整理下拉，只留選擇
        If Not flag_Enabled Then TIMS.GET_NewListItemVal(o_NewDDL1, v_NewData)
        '新的值為下拉物件 
        If flag_Enabled Then Common.SetListItem(o_NewDDL1, v_NewData)
        o_NewDDL1.Enabled = flag_Enabled 'False '鎖定 (已送出)
        Return v_OldData
    End Function

    ''' <summary>
    ''' 檢核dt, 產出正確的datarow
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <param name="dr"></param>
    ''' <param name="v_ChgItem"></param>
    Sub Utl_PlanReviseChkTableGetRow(ByRef dt As DataTable, ByRef dr As DataRow, ByVal v_ChgItem As String)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
        Else
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("PlanID") = rPlanID 'Request("PlanID")
            dr("ComIDNO") = rComIDNO 'Request("cid")
            dr("SeqNO") = rSeqNo 'Request("no")
            dr("SubSeqNo") = iSubSeqNO
            dr("CDate") = CDate(ApplyDate.Text) '.ToString("yyyy/MM/dd")
            dr("AltDataID") = TIMS.CINT1(v_ChgItem) 'ChgItem.SelectedValue
        End If
        Call UPDATE_PARTREDUC(dr)
    End Sub
End Class
