<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="menu.aspx.vb" Inherits="WDAIIP.menu" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href="./css/css.css" rel="stylesheet" type="text/css" />
    <link href="./css/style.css" rel="stylesheet" type="text/css" />
    <script language="javascript" type="text/javascript">
        function menustatas(flag) {
            if (flag == 'close') {
                //關閉
                //document.all.menushow.style.display = 'none';
                //document.all.menushow1.style.display = 'inline';
                document.getElementById("menushow").style.display = "none";
                document.getElementById("menushow1").style.display = "inline";
                //parent.document.all.MainBlock.cols = "30,*";
                parent.document.getElementById("MainBlock").cols = "30,*";
                if (parent.frames['mainFrame'].document.getElementById('FrameTable')) {
                    var intwid = '970px';
                    parent.frames['mainFrame'].document.getElementById('FrameTable').width = intwid;
                    parent.frames['titleFrame'].document.getElementById('tform1title1').style.backgroundImage = "url(./images/i2/title2.bmp)";
                    parent.frames['titleFrame'].document.getElementById('tform1title1').style.width = intwid;
                }
            }
            else if (flag == 'clear') {
                window.location.reload();
            }
            else {
                //開啟
                //document.all.menushow.style.display = 'inline';
                //document.all.menushow1.style.display = 'none';
                document.getElementById("menushow").style.display = "inline";
                document.getElementById("menushow1").style.display = "none";
                //parent.document.all.MainBlock.cols = "253,*";
                parent.document.getElementById("MainBlock").cols = "253,*";
                if (parent.frames['mainFrame'].document.getElementById('FrameTable')) {
                    parent.frames['mainFrame'].document.getElementById('FrameTable').width = '746';
                    parent.frames['titleFrame'].document.getElementById('tform1title1').style.backgroundImage = "url(./images/i2/title.bmp)";
                    parent.frames['titleFrame'].document.getElementById('tform1title1').style.width = "746px";
                }
            }
        }

        function NoAccount() {
            alert('職業訓練生活津貼管理系統,無此帳號!!');
            //return false;
        }

        //判斷所顯示清單內容
        function chkMenu(idx, cnt) {
            var imgOff;
            var imgOn;
            var labFunName;

            var obj;
            var imgOpen;
            var imgClose;
            var strDisplay = '';

            for (var i = 0; i <= 99; i++) {
                imgOff = document.getElementById('img' + i + '_off');
                imgOn = document.getElementById('img' + i + '_on');
                labFunName = document.getElementById('labFunName' + i);

                if (labFunName) {
                    if (i == idx) {
                        imgOff.style.display = 'none';
                        imgOn.style.display = 'inline';
                        labFunName.style.color = '#FF9900';
                        strDisplay = 'inline';

                    } else {
                        imgOff.style.display = 'inline';
                        imgOn.style.display = 'none';
                        labFunName.style.color = '#5288b2';
                        strDisplay = 'none';
                    }

                } else { break; }
            }
            return true;
        }
		
    </script>
</head>
<body style="margin-left: 0; margin-top: 0;">
    <noscript>
        你的瀏覽器不支援JavaScript!(請使用支持JavaScript的瀏覽器)</noscript>
    <form id="form1" method="post" runat="server">
    <div id="menushow1" style="display: none">
        <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <img src="./images/i2/button/left_open.bmp" alt="" style="cursor: hand" onclick="javascript:menustatas('open');" />
                </td>
            </tr>
            <tr>
                <td>
                    <img src="./images/i2/button/left_close.bmp" alt="" style="cursor: hand" onclick="javascript:menustatas('close');" />
                </td>
            </tr>
        </table>
    </div>
    <div id="menushow">
        <table cellspacing="0" cellpadding="0" border="0">
            <tr>
                <td class="fontMenu" valign="top">
                    <div id="div1" style="width: 200px;">
                        <asp:TreeView ID="TreeView1" runat="server">
                        </asp:TreeView>
                    </div>
                    <div id="div2">
                        <asp:Label ID="labmsg" runat="server" ForeColor="Red"></asp:Label>
                    </div>
                </td>
                <td valign="top" width="29">
                    <table id="Table1" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <img src="./images/i2/button/left_open.bmp" alt="" style="cursor: hand" onclick="javascript:menustatas('open');" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <img src="./images/i2/button/left_close.bmp" alt="" style="cursor: hand" onclick="javascript:menustatas('close');" />
                            </td>
                        </tr>
                        <tr>
                            <td>
                                <img src="./images/i2/button/left_clear.gif" alt="" style="cursor: hand" onclick="javascript:menustatas('clear');" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
