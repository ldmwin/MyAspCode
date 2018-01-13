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

Dim Column_Name,Column_Content,Parent_Name,Column_ID,Parent_ID


Function ColumnDetail()

Info_id=request.QueryString("id")

if Info_id="" or Info_id=0 then
	response.write "<script>alert('参数出错');history.go(-1);</Script>"
	response.end()
end if

Info_id=cint(Info_id)


sqlstr = "select c.*,(select Column_Name from Site_Columns where c.parent_id= id) as Parent_Name from Site_Columns c where c.status=1 and c.id = " & Info_id

Dim rs_Column : Set rs_Column = db.getRecordBySQL(sqlstr)

if rs_Column.eof or rs_Column.bof then
	response.write("数据查询失败")
	response.End()
end if

Column_Name = rs_Column("Column_Name")
Column_Content = rs_Column("Column_Content")
Parent_Name = rs_column("Parent_Name")
Parent_ID = rs_column("Parent_ID")

db.C(rs_Column)

ColumnDetail = Info_id

End Function

Column_ID =  ColumnDetail()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<meta name="author" content="#" /> 
	<meta name="robots" content="all" />
	<meta name="keywords" content="久友, 久友木门, 久友门业, 邢台久友, 邢台久友门业, 邢台久友木门, 邢台宏凯装饰, 河北久友, 河北久友门业, 久友官网" />
	<meta name="description" content="河北久友门业有限公司自成立以来，始终遵循“创新求变、团结拼搏、追求卓越”的企业精神，将“技术创新、人才建设、管理创新、市场营销、品牌塑造”作为工作重心，形成了“技术、人才、原材料、品牌、产品”五大竞争优势，使久友门业在激烈的市场竞争中，保持着强劲的发展势头。公司积极推行全面质量管理体系和全面预算管理体系，先后通过ISO9001:2008......" />
	<title>联系我们</title>
	<link rel="shortcut icon" href="../images/favicon.ico" type="image/x-icon" />
	<link href="../style/basic.css" rel="stylesheet" type="text/css" />
	<link href="../style/main.css" rel="stylesheet" type="text/css" />
	<link href="../style/contact.css" rel="stylesheet" type="text/css" />
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
			<!--#include file="../inc/contact_left_nav.asp"-->
<!--主体左侧结束-->

<!--主体右侧开始-->
			<div id="main_right">
				<div id="contact">
					<div id="contact_left">
						<div id="contact_left_title"><%=Column_Name%></div>
						<div id="contact_left_box">
							<%=Column_Content%>
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