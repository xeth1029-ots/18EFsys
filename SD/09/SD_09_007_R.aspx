<%@ Page AspCompat="true" Language="vb" AutoEventWireup="false" CodeBehind="SD_09_007_R.aspx.vb" Inherits="WDAIIP.SD_09_007_R" EnableEventValidation="false" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>學員退訓賠償未結案統計表</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1" />
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1" />
    <meta name="vs_defaultClientScript" content="JavaScript" />
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="<%=Request.ApplicationPath%>Scripts/jquery-3.7.1.min.js"></script>
    <%--<script type="text/javascript" src="../../js/selectControl.js.aspx" charset="UTF-8"></script>--%>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function print() {
            var msg = '';
            if (document.form1.syears.selectedIndex == 0) msg += '請選擇起始年度\n';
            if (document.form1.eyears.selectedIndex == 0) msg += '請選擇終至年度\n';
            if (document.form1.TPlan.selectedIndex == 0) msg += '請選擇訓練計畫\n';

            if (msg != '') {
                alert(msg);
                return false;
            }
        }
        function showFrame(arg) {
            document.getElementById('FrameObj').height = document.getElementById('HistoryRID').rows.length * 20;
            document.getElementById('FrameObj').style.display = document.getElementById('HistoryList2').style.display;
        }

        //        $(document).ready(function () {
        //            /* 起始下拉選單選項 */
        //            // 綁定 onchange
        //            $("select#TPlan").bind("change", LoadPlanList1);  // 綁定 onchange 到 LoadPlanList1()
        //        });

        //        function LoadPlanList1() {
        //            var sYearsList = document.getElementById('syears');
        //            var eYearsList = document.getElementById('eyears');
        //            var RIDValue = document.getElementById('RIDValue');
        //            var TPlanID = document.getElementById('TPlan');

        //            if (sYearsList.value != ''
        //                    && eYearsList.value != ''
        //                    && RIDValue.value != ''
        //                    && TPlanID.value != '') {
        //                var parms = "[";
        //                parms += "['sYears','" + sYearsList.value + "']";
        //                parms += ",['eYears','" + eYearsList.value + "']";
        //                parms += ",['RID','" + RIDValue.value + "']";
        //                parms += ",['TPlanID','" + TPlanID.value + "']";
        //                parms += "]";
        //                selectControl('ajaxClassList', 'ClassName', 'ClassCName', 'OCID', '請選擇', "", parms);
        //            }
        //            else {
        //                // 清空 ClassName
        //                var ClassName = document.getElementById("ClassName");
        //                if (ClassName) {
        //                    ClassName.length = 0;
        //                }
        //            }
        //        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
        <tr>
            <td>
                <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" class="font">
                    <tr>
                        <td>
                            首頁&gt;&gt;學員動態管理&gt;&gt;教務報表管理&gt;&gt;<font color="#990000">學員退訓賠償未結案清冊</font>
                        </td>
                    </tr>
                </table>
                <table id="Table3" class="table_sch" cellspacing="1" cellpadding="1">
                    <tr>
                        <td class="bluecol_need" width="100"> 年度區間 </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="syears" runat="server">
                            </asp:DropDownList>
                            年～
                            <asp:DropDownList ID="eyears" runat="server">
                            </asp:DropDownList>
                            年
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">
                            訓練機構
                        </td>
                        <td class="whitecol">
                            <asp:TextBox ID="center" runat="server" onfocus="this.blur()" Width="410px"></asp:TextBox>
                            <input type="button" value="..." id="Button2" name="Button2" runat="server" class="button_b_Mini" />
                            <input id="RIDValue" type="hidden" name="RIDValue" runat="server" /> 
                            <span id="HistoryList2" style="display: none; z-index: 1; position: absolute">
                                <asp:Table ID="HistoryRID" runat="server" Width="310px">
                                </asp:Table>
                            </span>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol">
                            <asp:Label ID="LabCJOB_UNKEY" runat="server">通俗職類</asp:Label>
                        </td>
                        <td colspan="3" class="whitecol">
                            <asp:TextBox ID="txtCJOB_NAME" runat="server" onfocus="this.blur()" Columns="30"></asp:TextBox>
                            <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                            <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol_need" width="100">
                            訓練計畫
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="TPlan" runat="server">
                            </asp:DropDownList>
                            <iframe id="FrameObj" style="display: none; left: 115px; position: absolute" scrolling="no" frameborder="0" width="310" height="50"></iframe>
                        </td>
                    </tr>
                    <tr>
                        <td class="bluecol" width="100">
                            班別
                        </td>
                        <td class="whitecol">
                            <asp:DropDownList ID="ClassName" runat="server">
                            </asp:DropDownList>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <br />
    <div style="width: 100%" align="center">
        <asp:Button ID="Button1" runat="server" Text="列印" CssClass="asp_Export_M"></asp:Button>
    </div>
    </form>
</body>
</html>
