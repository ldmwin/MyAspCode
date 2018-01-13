//function $(oid){
//		if(typeof(oid) == "string")
//		return document.getElementById(oid);
//		return oid;
//	}

function scrollDoor(){
}
scrollDoor.prototype = {
	sd : function(menus,divs,openClass,closeClass){
		var _this = this;
		if(menus.length != divs.length)
		{
			alert("菜单层数量和内容层数量不一样!");
			return false;
		}				
		for(var i = 0 ; i < menus.length ; i++)
		{	
			_this.$(menus[i]).value = i;
			//$(menus[i]).value = i;	
			//document.getElementById(menus[i]).value = i;
			_this.$(menus[i]).onmouseover = function(){
				//alert("ok");
				for(var j = 0 ; j < menus.length ; j++)
				{						
					_this.$(menus[j]).className = closeClass;
					_this.$(divs[j]).style.display = "none";
				}
				_this.$(menus[this.value]).className = openClass;	
				_this.$(divs[this.value]).style.display = "block";				
			}
		}
		},
	$ : function(oid){
		if(typeof(oid) == "string")
		return document.getElementById(oid);
		return oid;
	}
}
window.onload = function(){
	var SDmodel = new scrollDoor();
	//SDmodel.sd(["m01","m02","m03","m04"],["c01","c02","c03","c04"],"sd01","sd02");
	SDmodel.sd(["m01","m02"],["c01","c02"],"sd01","sd02");
	//var SDmodel1 = new scrollDoor();
	//SDmodel1.sd(["mm01","mm02","mm03","mm04"],["cc01","cc02","cc03","cc04"],"sdd01","sdd02");
	
}