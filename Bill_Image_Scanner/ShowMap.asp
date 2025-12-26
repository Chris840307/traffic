<HTML><HEAD><TITLE></TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<STYLE>BODY {
	PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px; FONT-FAMILY: 細明體
}
TD {
	PADDING-RIGHT: 0px; PADDING-LEFT: 0px; FONT-SIZE: 12px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px; FONT-FAMILY: 細明體
}
SPAN {
	POSITION: absolute
}
#map {
	LEFT: 0px; WIDTH: 800px; TOP: 0px; HEIGHT: 600px
}
#label {
	PADDING-RIGHT: 0px; PADDING-LEFT: 4px; Z-INDEX: 2; BACKGROUND: #ffe0e0; FILTER: alpha(opacity=80); LEFT: 0px; PADDING-BOTTOM: 0px; WIDTH: 80px; COLOR: #ee0000; PADDING-TOP: 3px; TOP: 0px; HEIGHT: 18px
}
#c_tl {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_tr {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_bl {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_br {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_tt {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_rr {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_bb {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
#c_ll {
	BORDER-RIGHT: medium none; BORDER-TOP: medium none; Z-INDEX: 1; BORDER-LEFT: medium none; BORDER-BOTTOM: medium none
}
</STYLE>


<!-- oncontextmenu=window.event.returnValue=false onselectstart="return false;" 
ondragstart="return false;" -->
<META content="MSHTML 6.00.3790.2541" name=GENERATOR></HEAD>
<BODY  bgColor=#d0d0d0 scroll=no>
<SPAN id=map>
<IMG onmouseup=FnMouseUp() onmousemove=FnMouseMove() onmousedown=FnMouseDown() onDblClick="ShowImgBig()" hideFocus src="<%=replace(trim(request("PicName")),"@2@","+")%>" border=0 name="theIMG" ID="theIMG" onLoad="getImg()"></SPAN>
<!--<SPAN id=c_tl><IMG id=arr_tl onmouseover="scroll_To( 10, 10,'tl')" onmouseout="scroll_Stop('tl')" height=21 
src="Pic/gis_arr_tl_0.gif" width=21></SPAN>
<!--<SPAN id=c_tr><IMG id=arr_tr onmouseover="scroll_To( 10,-10,'tr')" onmouseout="scroll_Stop('tr')" height=21 
src="Pic/gis_arr_tr_0.gif" width=21></SPAN>-->
<!--<SPAN id=c_bl><IMG id=arr_bl onmouseover="scroll_To(-10, 10,'bl')" onmouseout="scroll_Stop('bl')" height=21 
src="Pic/gis_arr_bl_0.gif" width=21></SPAN>
<!--<SPAN id=c_br><IMG id=arr_br onmouseover="scroll_To(-10,-10,'br')" onmouseout="scroll_Stop('br')" height=21 
src="Pic/gis_arr_br_0.gif" width=21></SPAN>-->
<SPAN id=c_tt><IMG id=arr_tt onmouseover="scroll_To( 10, 0,'tt')" onmouseout="scroll_Stop('tt')" height=21 
src="Pic/gis_arr_tt_0.jpg" width=21></SPAN>
<SPAN id=c_rr><IMG id=arr_rr onmouseover="scroll_To( 0,-10,'rr')" onmouseout="scroll_Stop('rr')" height=21 
src="Pic/gis_arr_rr_0.jpg" width=21></SPAN>
<SPAN id=c_bb><IMG id=arr_bb onmouseover="scroll_To(-10, 0,'bb')" onmouseout="scroll_Stop('bb')" height=21 
src="Pic/gis_arr_bb_0.jpg" width=21></SPAN>
<SPAN id=c_ll><IMG id=arr_ll onmouseover="scroll_To( 0, 10,'ll')" 
onmouseout="scroll_Stop('ll')" height=21 src="Pic/gis_arr_ll_0.jpg" width=21></SPAN> 
</BODY>
<SCRIPT language=JavaScript>

var stopimg=false;
var tMapWidth=document.theIMG.width, tMapHeight=document.theIMG.height; 			//image寬度,高度(影像大小)
var sMapWidth=document.theIMG.width, sMapHeight=document.theIMG.height;
var nMapWidth=screen.width-150,nMapHeight=screen.height-150;
var Xmin=0,Xmax=880;		//image影像真實座標
var Ymin=0,Ymax=620;

var tmpposX=0,tmpposY=0;						//image平移位置

var MouseX,MouseY;

var starX=0,starY=0;

var moveX=0,moveY=0;//位移量

var endX=0,endY=0;

var starMove=false;

var arr_tl_1=new Image();arr_tl_1.src="Pic/gis_arr_tl_1.gif";
var arr_tr_1=new Image();arr_tr_1.src="Pic/gis_arr_tr_1.gif";
var arr_bl_1=new Image();arr_bl_1.src="Pic/gis_arr_bl_1.gif";
var arr_br_1=new Image();arr_br_1.src="Pic/gis_arr_br_1.gif";
var arr_tt_1=new Image();arr_tt_1.src="Pic/gis_arr_tt_1.gif";
var arr_rr_1=new Image();arr_rr_1.src="Pic/gis_arr_rr_1.gif";
var arr_bb_1=new Image();arr_bb_1.src="Pic/gis_arr_bb_1.gif";
var arr_ll_1=new Image();arr_ll_1.src="Pic/gis_arr_ll_1.gif";

//再抓依次影像大小及初始化
function getImg(){
	sMapWidth=document.theIMG.width;
	sMapHeight=document.theIMG.height;
	document.theIMG.width=nMapWidth;
	document.theIMG.height=nMapHeight;
	tMapWidth=document.theIMG.width;
	tMapHeight=document.theIMG.height;
	initctrl();
}
function ShowImgBig(){
	if(nMapWidth==document.theIMG.width&&nMapHeight==document.theIMG.height){
		document.theIMG.width=sMapWidth;
		document.theIMG.height=sMapHeight;
		tMapWidth=sMapWidth;
		tMapHeight=sMapHeight;
	}else{
		document.theIMG.width=nMapWidth;
		document.theIMG.height=nMapHeight;
		tMapWidth=nMapWidth;
		tMapHeight=nMapHeight;
	}
	initctrl();
}
//------------------------------------------------------------------------------------------------
function FnFindXY(x,y){

var tXunit,tYunit;
tXunit=(Xmax-Xmin)/tMapWidth;
tYunit=(Ymax-Ymin)/tMapHeight;

var tFrameWidth,tFrameHeight;
tFrameWidth=document.body.clientWidth;		//Frame寬度
tFrameHeight=document.body.clientHeight;	//Frame高度

var tXO,tYO;
tXO=Xmin+tXunit*tFrameWidth/2;	//起始Frame中心點真實x座標(tmpposX=0)
tYO=Ymax-tYunit*tFrameHeight/2;	//起始Frame中心點真實y座標(tmpposY=0)

tmpX=tXO+Math.abs(tmpposX*tXunit);		//Frame中心點真實x座標
tmpY=tYO+tmpposY*tYunit;			//Frame中心點真實y座標

MouseX=formatnumber(tmpX+(x-(tFrameWidth/2))*tXunit,2)
MouseY=formatnumber(tmpY-(y-(tFrameHeight/2))*tYunit,2)
}


//------------------------------------------------------------------------------------------------
function formatnumber(numval,diginum) //調整數值格式
{
 var tstr;
 var tstr2;
 var i;
 tstr=Math.round(numval*Math.pow(10,diginum)).toString();
 if (tstr.length<=diginum){
  tstr2="0.";
  for (i=tstr.length;i<diginum;i++)
   tstr2+="0";
  tstr2+=tstr;}
 else{
  tstr2="."+tstr.substr(tstr.length-diginum,diginum);
  for (i=tstr.length-diginum;i>0;i-=3)
   if (i>3)
   		{tstr2= tstr.substr(i-3,3) + tstr2;}
   else
    	{tstr2=tstr.substr(0,i)+tstr2;} }
 return tstr2;
}


//------------------------------------------------------------------------------------------------
function scroll_To(y,x,icon){
	eval('arr_'+icon+'.src="Pic/gis_arr_'+icon+'_1.jpg"');

	if(tmpposY>=0&&y>=0)
		{y=0;}

	if(tmpposY<=(document.body.clientHeight-tMapHeight)&&y<=0)
		{y=0;}

	if(tmpposX>=0&&x>=0)
		{x=0;}

	if(tmpposX<=(document.body.clientWidth-tMapWidth)&&x<=0)
		{x=0;}

	if(!stopimg){
		tmpposX+=x;
		tmpposY+=y;
		map.style.posTop=tmpposY;
		map.style.posLeft=tmpposX;
		window.setTimeout("scroll_To("+y+","+x+",'"+icon+"')",1);}
}


//------------------------------------------------------------------------------------------------
function scroll_Stop(icon){
	stopimg=true;
	window.setTimeout("scroll_Ready('"+icon+"')",10);}

//------------------------------------------------------------------------------------------------
function scroll_Ready(icon){
	eval('arr_'+icon+'.src="Pic/gis_arr_'+icon+'_0.jpg"');
	stopimg=false;}

//------------------------------------------------------------------------------------------------
function initctrl(){

	//c_tl.style.posTop=20;
	//c_tl.style.posLeft=10;
	
	//c_tr.style.posTop=20;
	//c_tr.style.posLeft=document.body.clientWidth-31;
	
	//c_bl.style.posTop=document.body.clientHeight-31;
	//c_bl.style.posLeft=10;
	
	//c_br.style.posTop=document.body.clientHeight-31;
	//c_br.style.posLeft=document.body.clientWidth-31;
	
	c_tt.style.posTop=20;
	c_tt.style.posLeft=(document.body.clientWidth-21)/2;
	
	c_rr.style.posTop=(document.body.clientHeight-21)/2;
	c_rr.style.posLeft=document.body.clientWidth-31;
	
	c_bb.style.posTop=document.body.clientHeight-31;
	c_bb.style.posLeft=(document.body.clientWidth-21)/2;
	
	c_ll.style.posTop=(document.body.clientHeight-21)/2;
	c_ll.style.posLeft=10;
	
	var lastX,lastY;
	
	tXunit=(Xmax-Xmin)/tMapWidth;
	tYunit=(Ymax-Ymin)/tMapHeight;
	
	var JumpX,JumpY;
	var PixelX,PixelY;
	PixelX=eval((lastX-Xmin)/tXunit);
	PixelY=eval((Ymax-lastY)/tYunit);
	
	JumpX=PixelX-(document.body.clientWidth/2);
	JumpY=PixelY-(document.body.clientHeight/2);
	
	if (JumpX<0){JumpX=0;}
	if (JumpY<0){JumpY=0;}
	
			
	
	FirstJump();
	}

onresize=initctrl;


//------------------------------------------------------------------------------------------------
function FnMouseDown(){
		starMove=true;
		starX=event.x;
		starY=event.y;}
	

//------------------------------------------------------------------------------------------------
function FnMouseUp(){
	starMove=false;}


//------------------------------------------------------------------------------------------------
function FnMouseMove(){

	FnFindXY(event.x,event.y);
//	label.innerHTML="X:"+MouseX+"&nbsp;Y:"+MouseY;
	
	if (starMove==true){

		endX=event.x;
		endY=event.y;
		
		moveX=endX-starX;
		moveY=endY-starY;

		Jump_To(moveY,moveX);
		
		starX=endX;
		starY=endY;}
		
}

//------------------------------------------------------------------------------------------------
function Jump_To(y,x){
	newX=tmpposX+x
	newY=tmpposY+y

	if (newX<(document.body.clientWidth-tMapWidth))
		{newX=(document.body.clientWidth-tMapWidth);}

	if (newX>0)
		{newX=0;}

	if (newY<(document.body.clientHeight-tMapHeight))
		{newY=(document.body.clientHeight-tMapHeight);}

	if (newY>0)
		{newY=0;}

	tmpposX=newX;
	tmpposY=newY;

	map.style.posTop=Math.floor(tmpposY);
	map.style.posLeft=Math.floor(tmpposX );

}

//------------------------------------------------------------------------------------------------
function FirstJump(){

	//------------------------------
	/*var s, ss,sss,xx;
	var s = top.location.href   
	ss = s.split("?");
	sss=ss[1];
	xx=sss.split(",");
	*/
	var FirstGOx=440 ;  //xx[0];
	var FirstGOy=310  //xx[1];
	
	//alert(FirstGOx+" , "+FirstGOy)
	
	var tXunit,tYunit;
	var PixelX,PixelY;
	
	tXunit=(Xmax-Xmin)/tMapWidth;
	tYunit=(Ymax-Ymin)/tMapHeight;
	
	PixelX=eval(((FirstGOx-Xmin)/tXunit)-(document.body.clientWidth/2));
	PixelY=eval(((Ymax-FirstGOy)/tYunit)-(document.body.clientHeight/2));

	//alert(PixelX+" , "+PixelY)

	tmpposX=-PixelX;
	tmpposY=-PixelY;
	map.style.posTop=-PixelY;
	map.style.posLeft=-PixelX;

}

		</SCRIPT>
</HTML>
