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
	<meta name="keywords" content="长征驾校，邢台，邢台市长征机动车驾驶员培训学校，驾照" />
	<meta name="description" content="邢台市长征驾驶员培训学校,是一所以培训驾驶员为主,办理各种驾驶、行驶证件、年审、换证、申批等业务。本校师资力量雄厚，办学历史悠久，学校规模宏大，管理正规，曾连续几年被评为“优秀驾校”、“办学先进单位”、“信得过单位”。 本校现有从业多年的优秀教练数十名，“耐心教育，细心指导，把每位学员都培养成合格人才。”是本校的办学宗旨。为让学员更为深刻的学习和了解专业知识，最近学校又斥资数万元，首家引进最新模拟教学设备，相信，来本校学习会让你学习到最新、最专业的知识，长征驾校会让你踏遍理想的沃土，世界的角落。" />
	<title>邢台市长征驾驶员培训学校-长征驾校欢迎您!</title>
	<link rel="shortcut icon" href="../images/favicon.ico" type="image/x-icon" />
	<link href="../style/basic.css" rel="stylesheet" type="text/css" />
	<link href="../style/main.css" rel="stylesheet" type="text/css" />
	<link href="../style/news.css" rel="stylesheet" type="text/css" /></head>

<body>
	<!--顶部开始-->
		<!--#include file="../scontrol/top.asp"-->
	<!--顶部结束-->

<!--容器开始-->
	<div id="container">
		
		<div><img src="images/about_top_pic.jpg" width="1000" height="200" /></div>
		
<!--主体开始-->
		<div id="main_top">
			<div id="main_top_left">
				<div id="main_top_lefttitle"><%=Column_Name%></div>
				<div id="main_top_lefttitle2">inter</div>
			</div>
			<div id="main_top_right">
				<div id="main_topright_logo"><img src="../images/main_righttop_logo.jpg" /></div>
				<div class="main_righttop_title">
					<div class="main_righttop_titleleft"><strong><%=Column_Name%></strong>|</div>
					<div class="main_righttop_titleright"><img src="../images/main_leftnav_textbg.jpg" /></div>
					<div style="clear:both;"></div>
				</div>
			</div>
		</div>
			
		<div id="main">
			
			
			<div id="main_left">
				<div id="main_left_top">&nbsp;</div>
				<div id="main_leftnav_box">
					<%Function LeftSecNav(site_id,Parent_ID,menustr,path)%>
					<%
					Set rs_menutree = db.getRecordBySQL("select m.*,(select count(*) from Site_Columns s where s.Parent_id=m.id and s.status<>4) as childnodenum,(select Target from Sys_Target where m.Target = Sys_Target.id) as Link_Target,(select top 1 id from Site_Columns where parent_id = m.id and status=1 order by show_order desc,id desc) as childid from Site_Columns m where m.status<>4 and m.parent_id=" & Parent_ID & " and m.site_id=" & site_id & " order by m.show_order desc,m.id desc")
	
					do while not rs_menutree.eof
					
						dim target : target = rs_menutree("Link_Target") 

						menustr = menustr & "<div class=""main_leftnav_text"">"					

						
						'response.Write(str)
				'					
						if rs_menutree("column_type") = 1 then
							menustr = menustr & "<a href=""" & path & rs_menutree("Link") & """ target=""" & target & """>"						
													
						elseif rs_menutree("column_type") = 2 then					
							
							Set rs_link = db.getRecordBySQL("select top 1 * from Site_Column_Mould where status=1 and id=" & rs_menutree("mould_page") & " and (site_id=0 or site_id=" & site_id & ")order by show_order desc,id desc") 
							
							if not(rs_link.eof and rs_link.bof) then
								
								if rs_menutree("UseChildContent") = 0 then 
									menustr = menustr & "<a href=""" & path & rs_link("Mould_Path") & rs_link("Mould_File") & "?id=" & rs_menutree("id") &  """ target=""" & target & """>"
								else
									menustr = menustr & "<a href=""" & path & rs_link("Mould_Path") & rs_link("Mould_File") & "?id=" & rs_menutree("childid") &  """ target=""" & target & """>"
								
								end if
							else
								menustr = menustr & "<a href=""#"" target=""" & target & """>"
							end if
							
							db.C(rs_link)
						else
							menustr = menustr & "<a href=""#"" target=""" & target & """>"
						end if 
						 
						menustr = menustr & rs_menutree("column_name") & "</a></div>"
											
						rs_menutree.movenext()
						
					Loop
					
					db.C(rs_menutree)
					
					LeftSecNav = menustr
					
					%>
					<%End Function%>
					<%response.Write(LeftSecNav(Config.SiteID,Parent_ID,menustr,path))%>
				</div>
				<div style="height:50px;"></div>
				<div class="main_left_pic"><img src="images/registration_left_pic1.jpg" width="200" height="80" /></div>
				<div class="main_left_pic"><img src="images/registration_left_pic2.jpg" width="200" height="80" /></div>
				<div class="main_left_pic"><img src="images/registration_left_pic3.jpg" width="200" height="80" /></div>
				<!--<div><img src="../images/main_leftnav_bgbottom.jpg" /></div>-->
			</div>
			<div id="main_right">
				<div id="main_right_box">
					
					<div id="registration_box">
						
						<%=Column_Content%>
						
					</div>
					
				</div>
			</div>
			<div style="clear:both;"></div>
			
		<!--</div>-->
<!--主体结束-->
<!--尾部开始-->
		<!--#include file="../scontrol/footer.html"-->
</body>
</html>
<%
	set Config = nothing

	db.CloseConn()
	
	set db=nothing
%>