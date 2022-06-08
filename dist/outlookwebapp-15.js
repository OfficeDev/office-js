/* Outlook web application specific API library */
/* Version: 15.0.4420.1017 Build Time: 03/31/2014 */
/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

/*
	Copyright (c) Microsoft Corporation.  All rights reserved.
*/

/*
	Your use of this file is governed by the Microsoft Services Agreement http://go.microsoft.com/fwlink/?LinkId=266419.
*/

Type.registerNamespace("Microsoft.Office.WebExtension.MailboxEnums");Microsoft.Office.WebExtension.MailboxEnums.EntityType={MeetingSuggestion:"meetingSuggestion",TaskSuggestion:"taskSuggestion",Address:"address",EmailAddress:"emailAddress",Url:"url",PhoneNumber:"phoneNumber",Contact:"contact"};Microsoft.Office.WebExtension.MailboxEnums.ItemType={Message:"message",Appointment:"appointment"};Microsoft.Office.WebExtension.MailboxEnums.ResponseType={None:"none",Organizer:"organizer",Tentative:"tentative",Accepted:"accepted",Declined:"declined"};Microsoft.Office.WebExtension.MailboxEnums.RecipientType={Other:"other",DistributionList:"distributionList",User:"user",ExternalUser:"externalUser"};Microsoft.Office.WebExtension.MailboxEnums.AttachmentType={File:"file",Item:"item"}

Type.registerNamespace("OSF.DDA");OSF.DDA.OutlookAppOm=function(n,t,i){this.$$d_$1w_0=Function.createDelegate(this,this.$1w_0);this.$$d_$2Z_0=Function.createDelegate(this,this.$2Z_0);this.$$d_$2X_0=Function.createDelegate(this,this.$2X_0);this.$$d_$31_0=Function.createDelegate(this,this.$31_0);this.$$d_$2h_0=Function.createDelegate(this,this.$2h_0);this.$$d_$2e_0=Function.createDelegate(this,this.$2e_0);OSF.DDA.OutlookAppOm.$4=this;this.$L_0=n;this.$18_0=i;var u=this;var r=function(){i&&u.$7_0(1,"GetInitialData",null,u.$$d_$2e_0)};this.$19_0()?r():this.$36_0(r)};OSF.DDA.OutlookAppOm.$9=function(n,t,i,r){var f={};f[OSF.DDA.AsyncResultEnum.Properties.Value]=n;f[OSF.DDA.AsyncResultEnum.Properties.Context]=r;var u=null;if(0!==t){u={};u[OSF.DDA.AsyncResultEnum.ErrorProperties.Name]=t;u[OSF.DDA.AsyncResultEnum.ErrorProperties.Message]=i}return new OSF.DDA.AsyncResult(f,u)};OSF.DDA.OutlookAppOm.$D=function(n){if(!n)throw Error.create(_u.ExtensibilityStrings.l_ElevatedPermissionNeeded_Text);};OSF.DDA.OutlookAppOm.$C=function(n,t,i){if(n<t)throw Error.create(String.format(_u.ExtensibilityStrings.l_ElevatedPermissionNeededForMethod_Text,i));};OSF.DDA.OutlookAppOm.$H=function(n,t,i){if(Object.getType(n)!==t)throw Error.argumentType(i);};OSF.DDA.OutlookAppOm.$T=function(n,t,i,r){if(n<t||n>i)throw Error.argumentOutOfRange(r);};OSF.DDA.OutlookAppOm.$17=function(n,t,i,r){if(!$h.ScriptHelpers.isNullOrUndefined(n)){OSF.DDA.OutlookAppOm.$H(n,String,r);var u=n;OSF.DDA.OutlookAppOm.$T(u.length,t,i,r)}};OSF.DDA.OutlookAppOm.$21=function(n,t){var i=null;switch(n){case 1:case 2:case 12:case 3:break;case 4:var r=JSON.stringify(t.customProperties);i=[r];break;case 5:i=[t.body];break;case 8:case 9:i=[t.itemId];break;case 7:i=[OSF.DDA.OutlookAppOm.$e(t.requiredAttendees),OSF.DDA.OutlookAppOm.$e(t.optionalAttendees),t.start,t.end,t.location,OSF.DDA.OutlookAppOm.$e(t.resources),t.subject,t.body];break;case 11:case 10:i=[t.htmlBody];break;default:break}return i};OSF.DDA.OutlookAppOm.$e=function(n){return n?n.join(";"):null};OSF.DDA.OutlookAppOm.$1t=function(n,t){if($h.ScriptHelpers.isNullOrUndefined(n))return null;OSF.DDA.OutlookAppOm.$H(n,Array,t);var r=n;var u=null;var f=!1;OSF.DDA.OutlookAppOm.$T(r.length,0,OSF.DDA.OutlookAppOm.$1i,String.format("{0}.length",t));for(var e=0;e<r.length;e++)if($h.EmailAddressDetails.isInstanceOfType(r[e])){f=!0;break}f&&(u=[]);for(var i=0;i<r.length;i++)if(f){u[i]=$h.EmailAddressDetails.isInstanceOfType(r[i])?r[i].emailAddress:r[i];OSF.DDA.OutlookAppOm.$H(u[i],String,String.format("{0}[{1}]",t,i))}else OSF.DDA.OutlookAppOm.$H(r[i],String,String.format("{0}[{1}]",t,i));return u};OSF.DDA.OutlookAppOm.prototype={$3_0:null,$R_0:null,$1s_0:null,$1J_0:null,$L_0:null,$18_0:null,get_$O_0:function(){return this.$L_0.get_appName()},initialize:function(n){var t="itemType";this.$3_0=new $h.InitialData(n);1===n[t]?this.$R_0=new $h.Message(this.$3_0):3===n[t]?this.$R_0=new $h.MeetingRequest(this.$3_0):2===n[t]&&(this.$R_0=new $h.Appointment(this.$3_0));this.$1s_0=new $h.UserProfile(this.$3_0);this.$1J_0=new $h.Diagnostics(this.$3_0,this.$L_0.get_appName());$h.InitialData.$1(this,"item",this.$$d_$2h_0);$h.InitialData.$1(this,"userProfile",this.$$d_$31_0);$h.InitialData.$1(this,"diagnostics",this.$$d_$2X_0);OSF.DDA.OutlookAppOm.$4.get_$O_0()===64&&$h.InitialData.$1(this,"ewsUrl",this.$$d_$2Z_0)},makeEwsRequestAsync:function(n,t,i){if($h.ScriptHelpers.isNullOrUndefined(n))throw Error.argumentNull("data");if(n.length>OSF.DDA.OutlookAppOm.$1g)throw Error.argument("data",_u.ExtensibilityStrings.l_EwsRequestOversized_Text);OSF.DDA.OutlookAppOm.$C(this.$3_0.get_$8_0(),2,"makeEwsRequestAsync");var r=new $h.EwsRequest(i);var u=this;r.onreadystatechange=function(){4===r.get_$1l_1()&&t(r.$P_0)};r.send(n)},recordDataPoint:function(n){if($h.ScriptHelpers.isNullOrUndefined(n))throw Error.argumentNull("data");this.$7_0(0,"RecordDataPoint",n,null)},recordTrace:function(n){if($h.ScriptHelpers.isNullOrUndefined(n))throw Error.argumentNull("data");this.$7_0(0,"RecordTrace",n,null)},trackCtq:function(n){if($h.ScriptHelpers.isNullOrUndefined(n))throw Error.argumentNull("data");this.$7_0(0,"TrackCtq",n,null)},convertToLocalClientTime:function(n){var t=new Date(n.getTime());var i=t.getTimezoneOffset()*-1;if(this.$3_0&&this.$3_0.get_$13_0()){t.setUTCMinutes(t.getUTCMinutes()-i);i=this.$1P_0(t);t.setUTCMinutes(t.getUTCMinutes()+i)}var r=this.$g_0(t);r.timezoneOffset=i;return r},convertToUtcClientTime:function(n){var t=this.$2D_0(n);if(this.$3_0&&this.$3_0.get_$13_0()){var i=this.$1P_0(t);t.setUTCMinutes(t.getUTCMinutes()-i);i=n.timezoneOffset?n.timezoneOffset:t.getTimezoneOffset()*-1;t.setUTCMinutes(t.getUTCMinutes()+i)}return t},getUserIdentityTokenAsync:function(n,t){OSF.DDA.OutlookAppOm.$C(this.$3_0.get_$8_0(),1,"getUserIdentityTokenAsync");this.$1c_0(2,"GetUserIdentityToken",n,t)},getCallbackTokenAsync:function(n,t){OSF.DDA.OutlookAppOm.$C(this.$3_0.get_$8_0(),1,"getCallbackTokenAsync");if(64!==this.$L_0.get_appName())throw Error.notImplemented("The getCallbackTokenAsync is not supported by outlook for now.");this.$1c_0(12,"GetCallbackToken",n,t)},displayMessageForm:function(n){if($h.ScriptHelpers.isNullOrUndefined(n))throw Error.argumentNull("itemId");this.$7_0(8,"DisplayExistingMessageForm",{itemId:n},null)},displayAppointmentForm:function(n){if($h.ScriptHelpers.isNullOrUndefined(n))throw Error.argumentNull("itemId");this.$7_0(9,"DisplayExistingAppointmentForm",{itemId:n},null)},displayNewAppointmentForm:function(n){var u=OSF.DDA.OutlookAppOm.$1t(n.requiredAttendees,"requiredAttendees");var r=OSF.DDA.OutlookAppOm.$1t(n.optionalAttendees,"optionalAttendees");OSF.DDA.OutlookAppOm.$17(n.location,0,OSF.DDA.OutlookAppOm.$1h,"location");OSF.DDA.OutlookAppOm.$17(n.body,0,OSF.DDA.OutlookAppOm.$S,"body");OSF.DDA.OutlookAppOm.$17(n.subject,0,OSF.DDA.OutlookAppOm.$1j,"subject");if(!$h.ScriptHelpers.isNullOrUndefined(n.start)){OSF.DDA.OutlookAppOm.$H(n.start,Date,"start");var o=n.start;n.start=o.getTime();if(!$h.ScriptHelpers.isNullOrUndefined(n.end)){OSF.DDA.OutlookAppOm.$H(n.end,Date,"end");var i=n.end;if(i<o)throw Error.argumentOutOfRange("end",i,_u.ExtensibilityStrings.l_InvalidEventDates_Text);n.end=i.getTime()}}var t=null;if(u||r){t={};var s=n;for(var f in s){var e={key:f,value:s[f]};t[e.key]=e.value}u&&(t.requiredAttendees=u);r&&(t.optionalAttendees=r)}this.$7_0(7,"DisplayNewAppointmentForm",t||n,null)},$1M_0:function(n){$h.ScriptHelpers.isNullOrUndefined(n)||OSF.DDA.OutlookAppOm.$T(n.length,0,OSF.DDA.OutlookAppOm.$S,"htmlBody");this.$7_0(10,"DisplayReplyForm",{htmlBody:n},null)},$1L_0:function(n){$h.ScriptHelpers.isNullOrUndefined(n)||OSF.DDA.OutlookAppOm.$T(n.length,0,OSF.DDA.OutlookAppOm.$S,"htmlBody");this.$7_0(11,"DisplayReplyAllForm",{htmlBody:n},null)},$7_0:function(n,t,i,r){if(64===this.$L_0.get_appName())OSF._OfficeAppFactory.getClientEndPoint().invoke(t,r,i);else if(n){var u=OSF.DDA.OutlookAppOm.$21(n,i);var f=this;window.external.Execute(n,u,function(n,t){if(r){var u=n.getItem(0);var i=JSON.parse(u);r(t,i)}})}else r&&r(-2,null)},$2D_0:function(n){var t=new Date(n.year,n.month,n.date,n.hours,n.minutes,n.seconds,n.milliseconds?n.milliseconds:0);if(isNaN(t.getTime()))throw Error.format(_u.ExtensibilityStrings.l_InvalidDate_Text);return t},$g_0:function(n){var t={};t.month=n.getMonth();t.date=n.getDate();t.year=n.getFullYear();t.hours=n.getHours();t.minutes=n.getMinutes();t.seconds=n.getSeconds();t.milliseconds=n.getMilliseconds();return t},$2e_0:function(n,t){if(!n){this.initialize(t);this.displayName="mailbox";window.setTimeout(this.$$d_$1w_0,0)}},$1w_0:function(){this.$18_0()},$1c_0:function(n,t,i,r){if($h.ScriptHelpers.isNullOrUndefined(i))throw Error.argumentNull("callback");var u=this;this.$7_0(n,t,null,function(n,t){var u;if(n)u=OSF.DDA.OutlookAppOm.$9(null,1,String.format(_u.ExtensibilityStrings.l_InternalProtocolError_Text,n),r);else{var f=t;u=f.wasSuccessful?OSF.DDA.OutlookAppOm.$9(f.token,0,null,r):OSF.DDA.OutlookAppOm.$9(null,1,f.errorMessage,r)}i(u)})},$2h_0:function(){return this.$R_0},$31_0:function(){OSF.DDA.OutlookAppOm.$D(this.$3_0.get_$8_0());return this.$1s_0},$2X_0:function(){return this.$1J_0},$2Z_0:function(){OSF.DDA.OutlookAppOm.$D(this.$3_0.get_$8_0());return this.$3_0.get_$2F_0()},$1P_0:function(n){for(var r=this.$3_0.get_$13_0(),i=0;i<r.length;i++){var t=r[i];var f=parseInt(t.start);var u=parseInt(t.end);if(n.getTime()-f>=0&&n.getTime()-u<0)return parseInt(t.offset)}throw Error.format(_u.ExtensibilityStrings.l_InvalidDate_Text);},$19_0:function(){var n=!1;try{n=!$h.ScriptHelpers.isNullOrUndefined(_u.ExtensibilityStrings.l_EwsRequestOversized_Text)}catch(t){}return n},$36_0:function(n){for(var s=null,l="",a=document.getElementsByTagName("script"),o=a.length-1;o>=0;o--){var i=null;var v=a[o].attributes;if(v){var p=v.getNamedItem("src");p&&(i=p.value);if(i){var y=!1;i=i.toLowerCase();var f=i.indexOf("office_strings.js");if(f<0){f=i.indexOf("office_strings.debug.js");y=!0}if(f>0&&f<i.length){s=i.replace(y?"office_strings.debug.js":"office_strings.js","outlook_strings.js");var r=i.substring(0,f);var u=r.lastIndexOf("/",r.length-2);u===-1&&(u=r.lastIndexOf("\\",r.length-2));u!==-1&&r.length>u+1&&(l=r.substring(0,u+1));break}}}}if(s){var h=document.getElementsByTagName("head")[0];var t=null;var b=this;var e=function(){if(n&&(!t.readyState||t.readyState&&(t.readyState==="loaded"||t.readyState==="complete"))){t.onload=null;t.onreadystatechange=null;n()}};var c=this;var w=function(){if(!c.$19_0()){var n=l+"en-us/"+"outlook_strings.js";t.onload=null;t.onreadystatechange=null;t=c.$1G_0(n);t.onload=e;t.onreadystatechange=e;h.appendChild(t)}};t=this.$1G_0(s);t.onload=e;t.onreadystatechange=e;window.setTimeout(w,2e3);h.appendChild(t)}},$1G_0:function(n){var t=document.createElement("script");t.type="text/javascript";t.src=n;return t}};OSF.DDA.Settings=function(n){this.$u_0=n};OSF.DDA.Settings.$20=function(n){if(!n)return{};if(OSF.DDA.OutlookAppOm.$4.get_$O_0()===8){var t=n.SettingsKey;if(t)return OSF.DDA.SettingsManager.deserializeSettings(t)}return n};OSF.DDA.Settings.prototype={$u_0:null,$10_0:null,get_$J_0:function(){if(!this.$10_0){this.$10_0=OSF.DDA.Settings.$20(this.$u_0);this.$u_0=null}return this.$10_0},get:function(n){return this.get_$J_0()[n]},set:function(n,t){this.get_$J_0()[n]=t},remove:function(n){delete this.get_$J_0()[n]},saveAsync:function(){for(var n=[],i=0;i<arguments.length;++i)n[i]=arguments[i];var r=null;var u=null;if(n&&n.length>0){var t=n.length-1;if(Function.isInstanceOfType(n[t])){r=n[t];t--;t>=0&&(u=n[t].asyncContext)}}OSF.DDA.OutlookAppOm.$4.get_$O_0()===64?this.$3E_0(r,u):this.$3D_0(r,u)},$3D_0:function(n,t){var r=null;try{var f=OSF.DDA.SettingsManager.serializeSettings(this.get_$J_0());var e=JSON.stringify(f);var u={SettingsKey:e};OSF.DDA.RichClientSettingsManager.write(u)}catch(o){r=o}if(n){var i;i=r?OSF.DDA.OutlookAppOm.$9(null,1,r.message,t):OSF.DDA.OutlookAppOm.$9(null,0,null,t);n(i)}},$3E_0:function(n,t){var i=OSF.DDA.SettingsManager.serializeSettings(this.get_$J_0());var r=this;OSF._OfficeAppFactory.getClientEndPoint().invoke("saveSettingsAsync",function(i,r){if(n){var u;if(i)u=OSF.DDA.OutlookAppOm.$9(null,1,String.format(_u.ExtensibilityStrings.l_InternalProtocolError_Text,i),t);else{var f=r;u=f.error?OSF.DDA.OutlookAppOm.$9(null,1,f.errorMessage,t):OSF.DDA.OutlookAppOm.$9(null,0,null,t)}n(u)}},[i])}};Type.registerNamespace("$h");$h.Appointment=function(n){this.$$d_$2m_1=Function.createDelegate(this,this.$2m_1);this.$$d_$1W_1=Function.createDelegate(this,this.$1W_1);this.$$d_$q_1=Function.createDelegate(this,this.$q_1);this.$$d_$2t_1=Function.createDelegate(this,this.$2t_1);this.$$d_$1Z_1=Function.createDelegate(this,this.$1Z_1);this.$$d_$1X_1=Function.createDelegate(this,this.$1X_1);this.$$d_$n_1=Function.createDelegate(this,this.$n_1);this.$$d_$1T_1=Function.createDelegate(this,this.$1T_1);this.$$d_$1a_1=Function.createDelegate(this,this.$1a_1);$h.Appointment.initializeBase(this,[n]);$h.InitialData.$1(this,"start",this.$$d_$1a_1);$h.InitialData.$1(this,"end",this.$$d_$1T_1);$h.InitialData.$1(this,"location",this.$$d_$n_1);$h.InitialData.$1(this,"optionalAttendees",this.$$d_$1X_1);$h.InitialData.$1(this,"requiredAttendees",this.$$d_$1Z_1);$h.InitialData.$1(this,"resources",this.$$d_$2t_1);$h.InitialData.$1(this,"subject",this.$$d_$q_1);$h.InitialData.$1(this,"normalizedSubject",this.$$d_$1W_1);$h.InitialData.$1(this,"organizer",this.$$d_$2m_1)};$h.Appointment.prototype={getEntities:function(){return this.$0_0.$Q_0()},getEntitiesByType:function(n){return this.$0_0.$1U_0(n)},getRegExMatches:function(){OSF.DDA.OutlookAppOm.$C(this.$0_0.get_$8_0(),1,"getRegExMatches");return this.$0_0.$p_0()},getFilteredEntitiesByName:function(n){return this.$0_0.$m_0(n)},getRegExMatchesByName:function(n){OSF.DDA.OutlookAppOm.$C(this.$0_0.get_$8_0(),1,"getRegExMatchesByName");return this.$0_0.$1Y_0(n)},displayReplyForm:function(n){OSF.DDA.OutlookAppOm.$4.$1M_0(n)},displayReplyAllForm:function(n){OSF.DDA.OutlookAppOm.$4.$1L_0(n)},getItemType:function(){return Microsoft.Office.WebExtension.MailboxEnums.ItemType.Appointment},$1a_1:function(){return this.$0_0.get_$1m_0()},$1T_1:function(){return this.$0_0.get_$1N_0()},$n_1:function(){return this.$0_0.get_$1f_0()},$1X_1:function(){return this.$0_0.get_$a_0()},$1Z_1:function(){return this.$0_0.get_$14_0()},$2t_1:function(){return this.$0_0.get_$3C_0()},$q_1:function(){return this.$0_0.get_$1p_0()},$1W_1:function(){return this.$0_0.get_$1k_0()},$2m_1:function(){return this.$0_0.get_$37_0()}};$h.AttachmentDetails=function(n){this.$$d_$2g_0=Function.createDelegate(this,this.$2g_0);this.$$d_$2N_0=Function.createDelegate(this,this.$2N_0);this.$$d_$2v_0=Function.createDelegate(this,this.$2v_0);this.$$d_$2T_0=Function.createDelegate(this,this.$2T_0);this.$$d_$2l_0=Function.createDelegate(this,this.$2l_0);this.$$d_$2d_0=Function.createDelegate(this,this.$2d_0);this.$0_0=n;$h.InitialData.$1(this,"id",this.$$d_$2d_0);$h.InitialData.$1(this,"name",this.$$d_$2l_0);$h.InitialData.$1(this,"contentType",this.$$d_$2T_0);$h.InitialData.$1(this,"size",this.$$d_$2v_0);$h.InitialData.$1(this,"attachmentType",this.$$d_$2N_0);$h.InitialData.$1(this,"isInline",this.$$d_$2g_0)};$h.AttachmentDetails.prototype={$0_0:null,$2d_0:function(){return this.$0_0.id},$2l_0:function(){return this.$0_0.name},$2T_0:function(){return this.$0_0.contentType},$2v_0:function(){return this.$0_0.size},$2N_0:function(){var n=this.$0_0.attachmentType;return n<$h.AttachmentDetails.$Y.length?$h.AttachmentDetails.$Y[n]:Microsoft.Office.WebExtension.MailboxEnums.AttachmentType.File},$2g_0:function(){return this.$0_0.isInline}};$h.Contact=function(n){this.$$d_$2S_0=Function.createDelegate(this,this.$2S_0);this.$$d_$k_0=Function.createDelegate(this,this.$k_0);this.$$