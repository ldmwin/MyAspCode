<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Charset="utf-8"%>
<!--#include file="../Class/DBCtrl.asp" -->
<!--#include file="../Inc/Config.asp" -->
<!--#include file="../Inc/Function.asp" -->
<%
Dim Config: Set Config = New ClsConfig

Dim db : Set db = New DbCtrl
db.dbConnStr = Config.ConnStr(1,"../")
db.OpenConn()

Dim path : path ="../"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="author" content="#" /> 
	<meta name="robots" content="all" />
	<meta name="keywords" content="久友, 久友木门, 久友门业, 邢台久友, 邢台久友门业, 邢台久友木门, 邢台宏凯装饰, 河北久友, 河北久友门业, 久友官网" />
	<meta name="description" content="河北久友门业有限公司自成立以来，始终遵循“创新求变、团结拼搏、追求卓越”的企业精神，将“技术创新、人才建设、管理创新、市场营销、品牌塑造”作为工作重心，形成了“技术、人才、原材料、品牌、产品”五大竞争优势，使久友门业在激烈的市场竞争中，保持着强劲的发展势头。公司积极推行全面质量管理体系和全面预算管理体系，先后通过ISO9001:2008......" />
	<title>人才招聘</title>
	<link rel="shortcut icon" href="../images/favicon.ico" type="image/x-icon" />
	<link href="../style/basic.css" rel="stylesheet" type="text/css" />
	<link href="../style/main.css" rel="stylesheet" type="text/css" />
	<link href="../style/hr.css" rel="stylesheet" type="text/css" />
	<script language=javascript>
    function secBoard(n)
    {
    for(i=0;i<menu.childNodes.length;i++)
    menu.childNodes[i].className="sec1";
    menu.childNodes[n].className="sec2";
    for(i=0;i<bottom_nav_text.childNodes.length;i++)
    bottom_nav_text.childNodes[i].style.display="none";
    bottom_nav_text.childNodes[n].style.display="block";
    }
    </script>
</head>

<body>

<!--顶部开始-->
	<!--#include file="../inc/header.asp"-->
<!--顶部结束-->

<!--容器开始-->
	<div id="container">
<!--主体开始-->
		<div id="main">
		
<!--主体上部flash开始-->
			<!--#include file="../inc/main_boxtop.html"-->
<!--主体上部flash结束-->

<!--主体左侧开始-->
			<!--#include file="../inc/hr_left_nav.html"-->
<!--主体左侧结束-->

<!--主体右侧开始-->
			<div id="main_right">
				<div id="contact">
					<div id="contact_left">
						<div id="contact_left_title">企业简介 &nbsp; Contact</div>
						<div id="contact_left_box">
							<div class="contact_left_text">
								
								 <%Sub JobsList(InfoType,PageNum)%>
								<%
                                sqlstr="select * from [Jobs] where status=1 order by show_order desc,id desc"
                                                    
                                Dim rs_Jobs : Set rs_Jobs = db.getRecordBySQL(sqlstr)
                                
                                rs_Jobs.PageSize=PageNum
                            
                                pre = true
                                last = true
                                page = trim(Request.QueryString("page"))
                                  
                                if len(page) = 0 then
                                    intpage = 1
                                    pre = false
                                else
                                    if cint(page) =< 1 then
                                        intpage = 1
                                        pre = false
                                    else
                                        if cint(page) >= rs_Jobs.PageCount then
                                            intpage = rs_Jobs.PageCount
                                            last = false
                                        else
                                            intpage = cint(page)
                                        end if
                                    end if
                                end if
                                
                                if not rs_Jobs.eof then
                                    rs_Jobs.AbsolutePage = intpage
                                end if
                            
                                if not (rs_Jobs.eof or rs_Jobs.bof) then
                                    do until rs_Jobs.eof
                                %>

                                <table width="500" border="1">
								  <tr>
									<td>招聘岗位</td>
									<td>招聘人数</td>
									<td>发布日期</td>
								  </tr>
								  <tr>
									<td><a href="jobdetail.asp?id=<%=rs_Jobs("id")%>"><%=rs_Jobs("Job_Title")%></a></td>
									<td><%=rs_Jobs("Need_Num")%></td>
									<td>[<%response.Write(formatdatetime(rs_Jobs("add_Time"),2))%>]</td>
								  </tr>
								  <!--<tr>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
									<td>&nbsp;</td>
								  </tr>-->
								</table>
                                <!--综合新闻列表循环结束-->
                                <%
                                    rs_Jobs.movenext
                                    loop
                                else
                                    
                                %>
                                    没有数据
                                <%
                                end if											
                                %>
                                
                                <div id="page">
                                <%if rs_Jobs.pagecount > 0 then%>
                                    共<%=rs_Jobs.recordcount%>个职位&nbsp;当前页<%=intpage%>/<%=rs_Jobs.PageCount%>
                                <%else%>
                                    当前页0/0
                                <%end if%>
                                    <a href="index.asp?page=1">首页</a>| 
                                <%if pre then%>
                                    <a href="index.asp?page=<%=intpage -1%>">上页</a>| 
                                <%end if%>
                                <%if last then%>
                                    <a href="index.asp?page=<%=intpage +1%>">下页</a>| 
                                <%end if%>
                                    <a href="index.asp?page=<%=rs_Jobs.PageCount%>">尾页</a>|转到第 
                                    <select name="sel_page" onchange="javascript:location=this.options[this.selectedIndex].value;">
                                <%
                                for i = 1 to rs_Jobs.PageCount
                                if i = intpage then
                                %>
                                    <option value="index.asp?page=<%=i%>" selected><%=i%></option>
                                <%else%>
                                    <option value="index.asp?page=<%=i%>"><%=i%></option>
                                <%
                                end if
                                next
                                %>
                                </select>
                                页
                                </div>
                                <%
								db.C(rs_Jobs)
								%>
								<%end sub%>
                                
                                <%Call JobsList(0,10)%>	
							</div>
							<div class="contact_left_text">&nbsp;</div>

						</div>
					</div>
					<div id="contact_right">
						<div class="contact_right_pic"><img src="../images/05.jpg" width="125" height="75" /></div>
						<div class="contact_right_pic"><img src="../images/06.jpg" width="125" height="75" /></div>
						<div class="contact_right_pic"><img src="../images/15.jpg" width="125" height="75" /></div>
					</div>
					<div style="clear:both;"></div>
				</div>
			</div>	
<!--主体右侧结束-->

<!--主体底部开始-->

			<!--#include file="../inc/main_bottom.asp"-->

<!--主体底部结束-->
			
		</div>	
<!--主体结束-->

	</div>
<!--容器结束-->			

<!--尾部容器开始-->
	<!--#include file="../inc/footer.html"-->
<!--尾部容器结束-->
</body>
</html>
<%
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>