

<!--#include file="base.asp"-->
<%
'**********************************************
'下面两行是为了跳过微信的验证'
'response.write request("echostr")
'response.end
'**********************************************
dim signature        
dim timestamp       
dim nonce               
'dim echostr                '
dim signaturetmp

signature = Request("signature")
nonce = Request("nonce")
timestamp = Request("timestamp")
'**********************************************
dim ToUserName       
dim FromUserName
dim CreateTime      
dim MsgType               
dim Content   
dim EventKey    
dim Nickname     

set xml_dom = Server.CreateObject(xmldomString)
xml_dom.load request
FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text 
ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text 
MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text

if MsgType="text" then
	Content=xml_dom.getelementsbytagname("Content").item(0).text
elseif MsgType="event" then
	EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text
end if

if MsgType="event" then
        strEventType=xml_dom.getelementsbytagname("Event").item(0).text 
    if strEventType="CLICK" then
     	if check_empty_tag(fromusername)="false" then  '检查用户是否有标签，如果有标签才能查看报盘'
     		select case EventKey
     			case event_beef_future 
			  'strsend=get_resourcelist(FromUserName,ToUserName)
			  'strsend=bp_multi_items(FromUserName,ToUserName,"",news_beef_future)
			  		strsend=bp(FromUserName,ToUserName,title_beef_future,desc_beef_future,pic_beef,news_beef_future)
     			case event_beef_spot
			  		strsend=bp(FromUserName,ToUserName,title_beef_spot,desc_beef_spot,pic_beef,news_beef_spot)
			 	case event_pork_future
			  'strsend=bp(FromUserName,ToUserName,news_beef_future, eventkey)
			  		strsend=bp(FromUserName,ToUserName,title_pork_future,desc_pork_future,pic_pork,news_pork_future)
			 	case event_pork_spot
			  		strsend=bp(FromUserName,ToUserName,title_pork_spot,desc_pork_spot,pic_pork,news_pork_spot)
				case event_pork_product
			  'nickname=get_nickname(FromUserName)
			 'strsend=post_template(FromUserName,ToUserName,nickname)	
			  		strsend=bp(FromUserName,ToUserName,title_pork_product,desc_pork_product,pic_pork,news_pork_product)	 
				case event_beef_product	
			  		strsend=bp(FromUserName,ToUserName,title_beef_product,desc_beef_product,pic_beef,news_beef_product)	 			  		
			end select

		else  '没有标签返回提示信息'
			strsend=send_text(fromusername,tousername, msg_empty_tag)
		end if
	elseif strEventType="subscribe" then   '新用户关注时发送的消息'
		strsend=send_text(fromusername,tousername, msg_new_user)
    end if

Elseif msgtype="text" then
'strsend=text(fromusername,tousername,Content)
	if check_admin_tag(fromusername)="true" then    '检查是否在管理员组里，如果是才有权限输入特定代码'
		select case content
			case code_send_template   '群发模板消息'
				strsend=send_template_to_all_users(FromUserName,ToUserName)
			case code_close_bp	'关闭报盘'
				strsend=closebp(fromusername,tousername)
			case code_open_bp	'打开报盘'
				strsend=openbp(fromusername,tousername)
			case code_check_token	'检查token'
				strsend=send_text(fromusername,tousername,load_cache("get_token")&"//"&load_cache("token_date")&"//"&load_cache("expires_in"))
			case code_check_bpstatus	'检查报盘status'
				strsend=send_text(fromusername,tousername,load_cache("bp_status")&"//"&load_cache("bpstatus_date")&"//"&load_cache("bpstatus_changer"))
				'strsend=send_text(fromusername,ToUserName,load_cache_sysdb("bp_status")&"//"&load_cache_sysdb("bpstatus_date")&"//"&load_cache_sysdb("bpstatus_changer"))
			case code_get_news_list		'取得图文列表'
				strsend=get_resourcelist(fromusername,tousername)
			case code_test_mbxx			'单独发送模板消息'
				strsend=post_template(FromUserName,ToUserName,"")
			case code_get_tag_list		'取得所有标签列表'
				strsend=get_tag_list(fromusername,ToUserName)
			case code_get_self_tag		'取得自己的标签'
				strsend=get_self_tag(fromusername,ToUserName)	
			'case "userlist"
				'strsend=send_text(fromusername,ToUserName,get_all_users())
			case "black"    			'设置所有人为黑名单'
				strsend=set_black_list_by_tag(fromusername,tousername,0,"B")	
			case "unblack"    			'取消所有黑名单'
				strsend=set_black_list_by_tag(fromusername,tousername,0,"W")				
			case "settag"				'给黑名单上的人设置标签'
				strsend=set_tag_to_users(fromusername,tousername,tag_blacklist)	
			case "bltag"				'给标签组里的人设置为黑名单'
				strsend=set_black_list_by_tag(fromusername,tousername,tag_blacklist,"B")	
			case "wltag"				'取消标签组里人的黑名单'
				strsend=set_black_list_by_tag(fromusername,tousername,tag_blacklist,"W")						
			'case "black"    			'设置自己为黑名单'
				'strsend=set_black_list(fromusername,tousername,"")	
			'case "getbl"
				'strsend=get_black_list(fromusername,tousername)	
			case else
				'strsend=send_text(fromusername,tousername,Content)
				prestr=left(content,4)
				if prestr=code_set_template_title then
					strsend=set_template_title(fromusername,tousername,right(Content,len(content)-5))
				elseif prestr=code_set_template_brief then
					strsend=set_template_brief(fromusername,tousername,right(Content,len(content)-5))
				end if
		end select
	'else
		'strsend=send_text(fromusername,tousername,Content)
	end if

end if

if strsend<>"" then
response.write strsend
end if

set xml_dom=Nothing

'****************************************


%>
