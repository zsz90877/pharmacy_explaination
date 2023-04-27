<meta charset="UTF-8">
<meta name="viewport" content="height=device-height,width=device-width,initial-scale=1,maximum-scale=1,minimum-scale=1,user-scalable=no" />
<BODY style="background:url(bg.gif)">
<!--#include file="orazhiy.asp"-->
<!--#include file="sqlconn.asp"-->
<font style="font-size:24px"><%

if not ismobi() then
  response.write "请用手机查看。"
  response.end
end if

' 初始化变量
cfid=request("cfid")
sick_id=""
hzname=""
write_time=""
saymsg=""
t_fgnum=1


' 获取处方头
set rs=server.createobject("adodb.recordset")
set rs2=server.createobject("adodb.recordset")
rs.open "select sick_id,write_time,item_class from dispensary_prescribe_master where prescription_number='"&cfid&"'",oraconn,1,1
if rs.recordcount=0 then
  if cfid="" then
    response.write "处方号没有正确填写。"
  else
    response.write "没有找到该处方："&cfid
  end if
else
  if rs("item_class")<>"A" AND rs("item_class")<>"B" then
    response.write "处方类型不正确！<BR>目前只处理门诊药房、急诊药房、儿科药房、发热药房的处方。"
  else
    sick_id=rs("sick_id")
    write_time=rs("write_time")
    rs.close

    ' 获取患者信息
    rs.open "select name,(sysdate-birthdate)/365 as hzage,id_card_no,decode(sex,0,'男',1,'女','未知') hzsex from sick_basic_info where sick_id='"&sick_id&"'"
    hzsex=rs("hzsex")
    hzname=rs("name")
    hzage=rs("hzage")
    hzage=csng(hzage)

    if hzage>14 then
      if hzsex="男" then
        hzcall="先生"
      else
         if hzsex="女" then
           hzcall="女士"
         else
           hzcall="未知"
         end if
      end if
    else
      hzcall="小朋友"
    end if
    rs.close
    saymsg=saymsg&hzname&hzcall&"，您好！<br>这里是克拉玛依市中心医院药学部，药师团队为您服务！<br>"

    ' 获取处方细表内容（测试:105242049)
    sqlstr="select a.prescription_number,a.physic_code,a.physic_name,a.physic_spec,a.doseage,a.dose_unit,a.frequency,a.day_count,d.name fuse,"
    sqlstr=sqlstr&"decode(b.form,'2',a.doseage||a.dose_unit,'72',a.doseage||a.dose_unit,'67',a.doseage||a.dose_unit,decode(a.dose_unit,b.physic_unit,"
    sqlstr=sqlstr&"a.doseage||a.dose_unit,'丸',a.doseage||a.dose_unit,'掀',a.doseage||a.dose_unit,'吸',a.doseage||a.dose_unit,'IU',a.doseage||'单位',decode(a.doseage/b.min_dose,1.5,'一'||b.physic_unit||'半',2,"
    sqlstr=sqlstr&"'两'||b.physic_unit,0.5,'半'||b.physic_unit,round(a.doseage/b.min_dose,3)||b.physic_unit))) hscjl,"
    sqlstr=sqlstr&"a.quantity||a.physic_unit zsl,b.memo,b.min_dose,b.physic_unit,c.freq_memo from dispensary_prescribe_detail"
    sqlstr=sqlstr&" a,physic_dict_table b,PRESCRIBE_FREQUENCY_DICT c,base_dict d where a.prescription_number='"&cfid&"'"
    sqlstr=sqlstr&" and a.physic_code=b.physic_code and c.freq_describe=a.frequency and d.dict_name='TAKE_MEDICINE_WAYS_DICT' and d.code=a.usage"

    rs.open sqlstr,oraconn,3,2%>
    <table width=100% style="font-size:20px">
      <tr>
        <td>
          <font style='font-weight:bold;color:red;font-size:24px'><%=hzname%></font><%=hzcall%>，您好！</font><br>请核对您的药品。
        </td>
      </tr>
      <tr>
        <td bgcolor="#AAFFFF">处方内容</td>
      </tr>
      <tr>
        <td>
          处方号：<font style="color:#0000ff"><%=cfid%></font><br>
          处方日期：<font style="color:#0000ff"><%=write_time%></font><br>
          药品品种：<font style="color:#0000ff"><%=rs.recordcount%></font>个<p><%
          saymsg=saymsg&"您的这张处方上共有"&trim(cstr(rs.recordcount))&"种药品。<br>"
          i=0
          while not rs.eof
            i=i+1

            ' 显示给患者的信息处理
            response.write trim(cstr(i))&".<br><B>"&rs("physic_name")&"</B>，"&rs("zsl")&"<br>"
            response.write "规格:"&rs("physic_spec")&"<BR>"
            response.write "用法："&rs("fuse")&"，"&rs("freq_memo")&"，每次"
            if cdbl(rs("doseage"))<1 then
              response.write "0"&rs("doseage")
            else
              response.write rs("doseage")
            end if
            response.write rs("dose_unit")&"<BR>"

            ' 语音提示给患者的信息
            saymsg=saymsg&"第"&trim(cstr(i))&"个药是 "&rs("physic_name")&rs("zsl")&"，用法："&rs("fuse")&"，"&rs("freq_memo")&"，"
            saymsg=saymsg&"每次"
            if left(rs("hscjl"),1)="." then
              saymsg=saymsg&"0"&rs("hscjl")
            else
              saymsg=saymsg&rs("hscjl")
            end if
            rs2.open "SELECT a.csid,a.okflag,b.yyjytext,b.yyjysnd,b.yyjyvdo FROM dict_cross AS a FULL JOIN dict_ypbase AS b ON a.csmedb=b.ypbm WHERE a.cshosp='ZXYY' AND a.csmedh='"&trim(rs("physic_code"))&"'",sqlconn,1,1
            if rs2.recordcount>0 then
              if rs2("okflag")=1 then
                ypmemotxt=rs2("yyjytext")
                ypmemosnd=rs2("yyjysnd")
                ypmemovdo=rs2("yyjyvdo")
              else
                ypmemotxt=""
                ypmemosnd=""
                ypmemovdo=""
              end if
            else
              sqlconn.execute "insert into dict_cross (cshosp,csmedh,okflag) values ('ZXYY','"&rs("physic_code")&"',0)"
            end if
            rs2.close
            if ypmemotxt<>"" then
              if len(saymsg&"，要注意："&ypmemotxt&"。")>700 and i=3 then
                saymsg=saymsg&"[x2x]，要注意："&ypmemotxt&"。"
                t_fgnum=t_fgnum+1
              else
                saymsg=saymsg&"，要注意："&ypmemotxt&"。"
              end if
            else
              saymsg=saymsg+"。"
            end if
            saymsg=saymsg+"<br>"

            ' 用药交待信息是否显示给患者
            if ypmemotxt<>""  and 1=2  then
              response.write "注意事项："&ypmemotxt&"<P>"
            else
              response.write "<P>"
            end if

            rs.movenext
          wend
          saymsg=saymsg+"最后，祝您身体健康，早日康复！"

          rs.close%>
        </td>
      </tr>
    </table><%
  end if
end if

set rs2=nothing
set rs=nothing

saymsg=replace(saymsg,"▲","")
saymsg=replace(saymsg,"<br>","")
saymsg=replace(saymsg," ","")

saystr=split(saymsg,"[x2x]")
sayx=""
for i=0 to ubound(saystr)
  sayx=sayx&Server.URLEncode(saystr(i))&"[x2x]"
next
sayx=left(sayx,len(sayx)-5)

' 阅读用药注意事项
set tkhttp=server.createobject("MSXML2.XMLHTTP.3.0")
url="http://10.71.122.27/voice/ctos.asp?fnm="&cfid&"&saymsg="&sayx&"&tpx=1"

tkhttp.open "get",url,false
tkhttp.send ""
json=tkhttp.responseText
if right(json,2)="OK" then
  Set fso = Server.CreateObject("Scripting.FileSystemObject")
  for i=1 to t_fgnum
    call SaveRemoteFile("cfvoice/"&cfid&"_"&trim(i)&".mp3","http://10.71.122.27/voice/mp3/cf_"&Formatdate(date())&"_"&cfid&"_"&trim(i)&".mp3")
  next
  set fso=nothing
end if

response.write "<tr height=50><td>&nbsp;</td></tr><tr><td bgcolor='#AAFFFF'><font style='font-size:16px'><P>&nbsp;语音播报：</font></td></tr>"
  t_playstr=""""
  for i=t_fgnum to 1 step -1
    t_playstr=t_playstr+"http://med1.klmyszxyy.com/comm/cfvoice/"&cfid&"_"&trim(i)&".mp3"","""
  next
  t_playstr=left(t_playstr,len(t_playstr)-2)
' response.write t_playstr%>
  <div id="audioBox" height="400" width="100"></div>
  <script type="text/javascript">
    window.onload = function(){
      var arr = [<%=t_playstr%>];
      var myAudio = new Audio();
      myAudio.preload = true;
      myAudio.controls = true;
      myAudio.autostart = true;
      myAudio.src = arr.pop();
      myAudio.addEventListener('ended', playEndedHandler, false);
      myAudio.play();
      document.getElementById("audioBox").appendChild(myAudio);
      myAudio.loop = false;
      function playEndedHandler(){
        myAudio.src = arr.pop();
        myAudio.play();
        console.log(arr.length);
        !arr.length && myAudio.removeEventListener('ended',playEndedHandler,false);
      }
    }
  </script><%
response.write "<tr><td>"&sayword+"</td></tr></table>"

' 程序结束，显示版权页
response.write "<P>&nbsp;<font style='font-size:12px'><hr align=left width='95%'>克拉玛依市中心医院 @Zxyy.PH. Y V2.1.210127<br>End...<P>"
saymsg=replace(saymsg,"[x2x]","")
'response.write "<P><B>语音测试：</B><br>"&saymsg

response.end



' 各种内部自定义函数
Function parseJSON(str) 
  ' json变量处理
  If Not IsObject(scriptCtrl) Then 
    Set scriptCtrl = Server.CreateObject("MSScriptControl.ScriptControl") 
    scriptCtrl.Language = "JScript" 
    scriptCtrl.AddCode "Array.prototype.get = function(x) { return this[x]; }; var result = null;" 
  End If 
  scriptCtrl.ExecuteStatement "result = " & str & ";" 
  Set parseJSON = scriptCtrl.CodeObject.result 
End Function 

sub SaveRemoteFile(LocalFileName,RemoteFileUrl)
  dim Ads,Retrieval,GetRemoteData
  Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
  With Retrieval
    .Open "Get", RemoteFileUrl, False, "", ""
    .Send
    GetRemoteData = .ResponseBody
  End With
  Set Retrieval = Nothing
  Set Ads = Server.CreateObject("Adodb.Stream")
  With Ads
    .Type = 1
    .Open
    .Write GetRemoteData
    .SaveToFile server.MapPath(LocalFileName),2
    .Cancel()
    .Close()
  End With
  Set Ads=nothing
end sub

Function Formatdate(thedate)
  dYear = Year(thedate)
  dMonth = Month(thedate)
  dDay = Day(thedate)
  Formatdate = dYear&Right("0"&dMonth,"2")&Right("0"&dDay,"2")
End Function


function ismobi()
  HTTP_ACCEPT=Request.ServerVariables("HTTP_ACCEPT")
  HTTP_USER_AGENT=LCase(Request.ServerVariables("HTTP_USER_AGENT"))
  HTTP_X_WAP_PROFILE=Request.ServerVariables("HTTP_X_WAP_PROFILE")
  HTTP_UA_OS=Request.ServerVariables("HTTP_UA_OS")
  HTTP_VIA=LCase(Request.ServerVariables("HTTP_VIA"))
  ismobi=False
  If instr(HTTP_ACCEPT,"vnd.wap")>0 Then
    ismobi=True
  elseIf HTTP_USER_AGENT="" Then
    ismobi=True
  elseIf HTTP_X_WAP_PROFILE<>"" Then
    ismobi=True
  elseIf HTTP_UA_OS<>"" Then
    ismobi=True
  elseif instr(HTTP_VIA,"wap")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"netfront")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"iphone")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"opera mini")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"ucweb")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"windows ce")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"symbianos")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"java")>0 Then
    ismobi=True
  elseif instr(HTTP_USER_AGENT,"android")>0 Then
    ismobi=True
  End if
end function

%>
