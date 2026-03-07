<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="checkie.aspx.vb" Inherits="WDAIIP.checkie" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
	<title></title>
</head>
<body>
	<form id="form1" runat="server">
	<script type="text/javascript" src="js/checkIE.js"></script>
	<script type="text/javascript">
		if (window.addEventListener) {
			window.addEventListener('load', showEachClientCheckResult, false);

		} else if (window.attachEvent) {
			window.attachEvent('onload', showEachClientCheckResult);

		}
		var MsgOK = '<font color="green"><b>成功！</b></font>';
		var MsgNG = '<font color="red"><b>失敗！</b></font>';

		var Msg = new Array(
            new Array(
                MsgNG + '你正在使用的瀏覽器不被支援。<a href="#browser">點選這裡</a>'
                , ''
                , MsgOK + '你正在使用正確的瀏覽器。'
            )
            ,
            new Array(
                ''
                , ''
                , MsgOK + '你的瀏覽器允許JavaScript。'
            )
            ,
            new Array(
                MsgNG + '你的瀏覽器不允許Cookies。<a href="#cookies">點選這裡</a>'
                , ''
                , MsgOK + '你的瀏覽器允許Cookies。'
            )
            ,
            new Array(
                MsgNG + '你的瀏覽器不允許彈出視窗。<a href="#popup">點選這裡</a>'
                , ''
                , MsgOK + '你的瀏覽器允許彈出視窗。'
            )
            ,
            new Array(
                MsgNG + '你需要安裝Flash播放器。<a href="#flash">點選這裡</a>'
                , MsgNG + '你需要更新你的Flash播放器。<a href="#flash">點選這裡</a>'
                , MsgOK + '你的瀏覽器正在使用正確的Flash播放器。'
            )
        )

	</script>
	<div>
		<%--<div class="breadcrumbs">
            <a href='<%=Page.ResolveUrl("~")+"internet/index/index.aspx" %>'>首頁</a> 
            <a href='<%=Page.ResolveUrl("~")+"internet/index/List.aspx?uid=858&pid=1684" %>' title="其他">其他</a> 
            瀏覽器設定檢查 
        </div>--%>
		<div class="pgmainhd">
			<div class="icon">
				<a id="A3" title="返回" class="l" href="#" onclick="window.open('MOICA_Login.aspx','_self');">
					<img alt="返回" src="images/newsicon2.gif"></a>
			</div>
			瀏覽器設定檢查
		</div>
		<table width="100%" cellspacing="0" cellpadding="0" border="0">
			<tbody>
				<tr>
					<td valign="top">
						<table width="100%" cellspacing="0" cellpadding="0" border="0">
							<tbody>
								<tr>
									<td>
										<table width="96%" cellspacing="0" cellpadding="0" border="0" align="center">
											<tbody>
												<tr>
													<td>
														<b>■為了可以正常使用這個系統，你的瀏覽器需要：</b>
													</td>
												</tr>
												<tr>
													<td>
														&nbsp;
													</td>
												</tr>
												<tr>
													<td>
														<img src="images/IconList001.jpg" alt="*" />
														<a href="#browser" style="color: #0066ff">Microsoft Internet Explorer 9.0 或以上版本</a><br />
														<img src="images/IconList001.jpg" alt="*" />
														<a href="#JavaScript" style="color: #0066ff">允許 JavaScript</a><br />
														<img src="images/IconList001.jpg" alt="*" />
														<a href="#cookies" style="color: #0066ff">允許 Cookies</a><br />
														<img src="images/IconList001.jpg" alt="*" />
														<a href="#popup" style="color: #0066ff">允許彈出視窗</a><br />
														<img src="images/IconList001.jpg" alt="*" />
														<a href="#flash" style="color: #0066ff">安裝有 Macromedia Flash 播放器 7.0 或以上版本</a><br />
														&nbsp;
													</td>
												</tr>
												<tr>
													<td>
														<table width="95%" cellspacing="0" cellpadding="20" bordercolor="#74bfd5" border="1" align="center">
															<tbody>
																<tr>
																	<td bgcolor="#e8faff">
																		我們正在檢查你的瀏覽器以確定是否符合以上要求。如果您看到任何一項<span>失敗！</span>，請按照提示操作。如果每一項都是<span class="closereason_text01">成功！</span>，你可以直接開始使用此系統。
																		<br />
																		<br />
																		<table width="95%">
																			<tr>
																				<td width="3%">
																				</td>
																				<td>
																					瀏覽器：
																				</td>
																				<td>
																					<span id="browser_check">檢查中...</span>
																				</td>
																			</tr>
																			<tr>
																				<td width="3%">
																				</td>
																				<td>
																					JavaScript：
																				</td>
																				<td>
																					<span id="js_check">檢查中...</span><noscript><font color="red"><b> 失敗！</b></font>你的瀏覽器不允許 JavaScript。<a href="#JavaScript">點選這裡 </a>
																					</noscript>
																				</td>
																			</tr>
																			<tr>
																				<td width="3%">
																				</td>
																				<td>
																					Cookies：
																				</td>
																				<td>
																					<span id="cookie_check">檢查中...</span>
																				</td>
																			</tr>
																			<tr>
																				<td width="3%">
																				</td>
																				<td>
																					彈出視窗：
																				</td>
																				<td>
																					<span id="popup_check">檢查中...</span>
																				</td>
																			</tr>
																			<tr>
																				<td width="3%">
																				</td>
																				<td>
																					Flash播放器：
																				</td>
																				<td>
																					<span id="flash_check">檢查中...</span>
																				</td>
																			</tr>
																		</table>
																	</td>
																</tr>
															</tbody>
														</table>
													</td>
												</tr>
												<tr>
													<td>
														&nbsp;
													</td>
												</tr>
												<tr>
													<td>
														如果上面有任何一項<span style="color: #ff0000">失敗！</span>，請看以下部分。
													</td>
												</tr>
												<tr>
													<td>
														&nbsp;
													</td>
												</tr>
												<tr>
													<td>
														&nbsp;
													</td>
												</tr>
												<tr>
													<td>
														&nbsp;
													</td>
												</tr>
												<tr>
													<td>
														<table width="100%" cellspacing="0" cellpadding="0" border="0">
															<tbody>
																<tr>
																	<td width="2%" valign="top" nowrap="nowrap">
																		<b>1.</b>
																	</td>
																	<td>
																		<b><a name="browser">瀏覽器</a></b><br />
																		請採用 Internet Explorer 瀏覽此系統。<br />
																		如果你使用的是舊版本的 Internet Explorer，請到<a href="http://windowsupdate.microsoft.com/" target="_blank" style="color: #0066ff">Windows Update</a>更新。<br />
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td style="padding: 0.7em 0; text-align: center; border-top: 1px dotted #cccccc; border-bottom: 1px dotted #cccccc; margin: 0.5em 0;">
																		<a href="javascript:window.location.reload();">
																			<img src="images/ck_01.gif" alt="再次檢查" /></a>
																	</td>
																</tr>
																<tr>
																	<td colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top" colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		<b>2.</b>
																	</td>
																	<td>
																		<b><a name="JavaScript">JavaScript</a></b><br />
																		如果被禁用了，請根據以下步驟啟用：
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td>
																		1. 打開瀏覽器的<span> 「工具」－「網際網路選項」－「安全性」－「網際網路」－「自訂層級」</span><br />
																		2. 找到「指令碼處理」部分，如下圖<br />
																		<img src="images/chk_env01.jpg" alt="安全性設定" /><br />
																		<br />
																		3. 確定「Active scripting」是「啟用」狀態。<br />
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td style="padding: 0.7em 0; text-align: center; border-top: 1px dotted #cccccc; border-bottom: 1px dotted #cccccc; margin: 0.5em 0;">
																		<a href="javascript:window.location.reload();">
																			<img src="images/ck_01.gif" alt="再次檢查" /></a>
																	</td>
																</tr>
																<tr>
																	<td colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top" colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		<b>3.</b>
																	</td>
																	<td>
																		<b><a name="cookies">Cookies</a></b><br />
																		如果被禁用了，請根據以下步驟啟用：
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td>
																		1. 打開瀏覽器的 <span>「工具」－「網際網路選項」－「隱私」－「進階」</span><br />
																		2. 確定「覆寫自動 cookie 處理」是<strong>沒有勾選</strong>，如下圖<br />
																		<img src="images/chk_env02.jpg" alt="進階隱私設定" /><br />
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td style="padding: 0.7em 0; text-align: center; border-top: 1px dotted #cccccc; border-bottom: 1px dotted #cccccc; margin: 0.5em 0;">
																		<a href="javascript:window.location.reload();">
																			<img src="images/ck_01.gif" alt="再次檢查" /></a>
																	</td>
																</tr>
																<tr>
																	<td colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top" colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		<b>4.</b>
																	</td>
																	<td>
																		<b><a name="popup">彈出視窗</a></b><br />
																		如果你的瀏覽器禁止任何彈出視窗，請把它取消。或者你安裝了任何可以攔截彈出視窗的工具，也請把該功能也取消。
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td>
																		1. Windows XP 作業系統預設已安裝 SP2 修補更新程式，請修改下列兩項設定：<br />
																		請檢查 Internet Explorer : <span>「工具」－「快顯封鎖程式」－「關閉快顯封鎖程式」</span>；將快顯封鎖程式關閉。<br />
																		<img src="images/chk_env03.jpg" alt="關閉快顯封鎖程式" /><br />
																		請檢查 Internet Explorer : <span>「工具」－「網際網路選項」－「安全性」－「自訂層級」－「下載簽名的ActiveX」</span>選擇「提示」; 「自動提示ActiveX控制項」，選擇「啟用」。按「確定」鈕後，重新開啟 Internet Explorer 瀏覽器。<br />
																		<img src="images/chk_env04.jpg" alt="自動提示ActiveX控制項" /><br />
																		<br />
																		2. Yahoo的工具列中的「選項」，<strong>取消</strong>「Enable Pop-UP blocker」<br />
																		<img src="images/chk_env05.jpg" alt="取消「Enable Pop-UP blocker」" /><br />
																		<br />
																		3. Google的工具列中的「選項」，取消「彈出視窗攔截器」功能。<br />
																		<img src="images/chk_env06.jpg" alt="彈出視窗攔截器" /><br />
																		<br />
																		如果是以下情況，代表目前是允許彈出視窗<br />
																		<img src="images/chk_env07.jpg" alt="允許彈出視窗" /><br />
																		<br />
																		4. MSN的工具列中的「選項」，取消「封鎖快顯視窗」功能。<br />
																		<img src="images/chk_env08.jpg" alt="封鎖快顯視窗" /><br />
																		<br />
																		5. 使用者，請選擇「允許 tims.etraining.gov.tw 使用彈出型視窗」。<br />
																		<img src="images/chk_env09.jpg" alt="允許www.taiwanjobs.gov.tw使用彈出型視窗" /><br />
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td style="padding: 0.7em 0; text-align: center; border-top: 1px dotted #cccccc; border-bottom: 1px dotted #cccccc; margin: 0.5em 0;">
																		<a href="javascript:window.location.reload();">
																			<img src="images/ck_01.gif" alt="再次檢查" /></a>
																	</td>
																</tr>
																<tr>
																	<td valign="top" colspan="2">
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		<b>5.</b>
																	</td>
																	<td>
																		<b><a name="flash">Flash播放器</a></b><br />
																		如果你使用的是舊版本的Flash播放器，或者是沒有Flash播放器，請到<a href="http://www.macromedia.com/go/getflashplayer" target="_blank" style="color: #0066ff">Macromedia Download Site</a>下載並安裝。<br />
																		&nbsp;
																	</td>
																</tr>
																<tr>
																	<td valign="top">
																		&nbsp;
																	</td>
																	<td style="padding: 0.7em 0; text-align: center; border-top: 1px dotted #cccccc; border-bottom: 1px dotted #cccccc; margin: 0.5em 0;">
																		<a href="javascript:window.location.reload();">
																			<img src="images/ck_01.gif" alt="再次檢查" /></a>
																	</td>
																</tr>
															</tbody>
														</table>
													</td>
												</tr>
											</tbody>
										</table>
									</td>
								</tr>
							</tbody>
						</table>
					</td>
				</tr>
			</tbody>
		</table>
	</div>
	</form>
</body>
</html>
