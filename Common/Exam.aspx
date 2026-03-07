<%@ OutputCache Location="None" %>

<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Exam.aspx.vb" Inherits="WDAIIP.Exam" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html lang="zh-TW">
<head>
    <title>科系所代碼查詢</title>
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

        function return_value(sub_id, sub_name) {
            var myfield;
            myfield = getParamValue("exam_id_field");
            if (myfield != "") {
                opener.document.getElementById(myfield).value = sub_id;
            }
            myfield = getParamValue("exam_name_field");
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
        <table summary="科系代碼分類大項" id="MainBlock" bordercolor="#0000cc" width="90%" align="center" border="2" runat="server">
            <tr>
                <td>
                    <table summary="科系代碼資料區" width="100%" align="center">
                        <tr>
                            <td align="center"><strong><font style="font-family: 標楷體; color: #000066; font-size: 18px"><strong><strong><font face="標楷體" color="#000066" size="4">技 能 檢 定 名 稱</font></strong></strong></font></strong></td>
                        </tr>
                        <tr>
                            <td align="center"><font style="color: #ff0000; font-size: 14px"><font color="#ff0000" size="2">※請選擇技能檢定名稱 ※</font></font></td>
                        </tr>
                    </table>
                    <table summary="科系代碼資料區" id="MainList" cellspacing="2" cellpadding="6" width="100%" align="center" border="0" runat="server">
                    </table>
                    <div align="center">【<a title="清除科系所代碼" href="javascript:return_value('','','','');">取消</a>】</div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>