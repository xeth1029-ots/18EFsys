<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="JobCode.aspx.vb" Inherits="WDAIIP.JobCode" %>

<html>
<head>
    <title>標準職業類別查詢</title>
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <meta http-equiv="Content-Language" content="zh-tw">
    <meta http-equiv="Content-Type" content="text/html; charset=big5">
    <style type="text/css">
        A { font-size: 10pt; font-family: "新細明體"; }
            A:visited { color: #0000ff; }
            A:link { color: #0000ff; }
            A:hover { color: #9933cc; }
            A:active { color: #9933cc; }
    </style>
    <script type="text/javascript" language="javascript">
        function getParamValue(name) {
            var querystring;
            var values;
            var result = "";
            if (location.search.length > 1) {
                querystring = unescape(location.search + "&");
                var re = new RegExp("[\?|\&]" + name + "=(.+)\&");
                values = querystring.match(re);
                if (values != null) {
                    result = values[1];
                    if (result.indexOf("&") != -1) {
                        result = result.substring(0, result.indexOf("&"));
                    }
                }
            }
            return result;
        }

        function showDetailTable(myname, myid) {
            document.getElementById("todo").value = "QUERY";
            document.getElementById("main_name").value = myname;
            document.getElementById("main_id").value = myid;
            document.form1.submit();
        }

        function return_value(main_id, main_name, sub_id, sub_name) {
            var myfield;
            myfield = getParamValue("job_id_field");
            if (myfield != "") {
                opener.document.getElementById(myfield).value = sub_id;
            }
            myfield = getParamValue("job_name_field");
            if (myfield != "") {
                opener.document.getElementById(myfield).value = sub_name;
            }
            window.close();
        }
    </script>
</head>
<body bgcolor="#e6efff">
    <form id="form1" method="post" runat="server">
        <input id="todo" type="hidden" runat="server" name="todo">
        <input id="main_id" type="hidden" runat="server" name="main_id">
        <input id="main_name" type="hidden" runat="server" name="main_name">
        <input id="sub_id" type="hidden" runat="server" name="sub_id">
        <input id="sub_name" type="hidden" runat="server" name="sub_name">
        <table id="MainBlock" bordercolor="#0000cc" width="90%" align="center" border="2" runat="server">
            <tr>
                <td>
                    <table width="100%" align="center">
                        <tr>
                            <td align="center"><strong><font face="標楷體" color="#000066" size="4">工 作 類 別</font></strong></td>
                        </tr>
                        <tr>
                            <td align="center"><font color="#ff0000" size="2">※請按下工作職業分類名稱後選擇下一層分類選單※</font></td>
                        </tr>
                    </table>
                    <table id="MainList" cellspacing="2" cellpadding="6" width="100%" align="center" border="0" runat="server">
                    </table>
                    <div align="center">【<a title="關閉視窗" href="javascript:window.close();">關閉</a>】&nbsp; 【<a title="清除工作職業分類" href="javascript:return_value('','','','');">清除</a>】</div>
                </td>
            </tr>
        </table>
        <table id="SubBlock" bordercolor="#0000cc" width="98%" align="center" border="2" runat="server">
            <tr>
                <td>
                    <table width="100%" align="center">
                        <tr>
                            <td align="center"><strong><font face="標楷體" color="#000066" size="4"><strong><font face="標楷體" color="#000066" size="4"><strong><font face="標楷體" color="#000066" size="4">工 作 類 別</font></strong>--</font></strong><asp:Label ID="lblMainName" runat="server"></asp:Label></font></strong></td>
                        </tr>
                        <tr>
                            <td align="center"><font color="#ff0000" size="2"><font color="#ff0000" size="2">※請選擇職業名稱※</font></font></td>
                        </tr>
                    </table>
                    <table id="SubList" cellspacing="2" cellpadding="6" width="100%" align="center" border="0" runat="server">
                    </table>
                    <div align="center">【<a title="關閉視窗" href="javascript:window.close();">關閉</a>】&nbsp; 【<a title="清除縣市鄉鎮" href="javascript:return_value('','','','');">清除</a>】&nbsp;【<a title="返回上一層選單" href="javascript:history.go(-1);">回上頁</a>】</div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>