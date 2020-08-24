
<%

				vbody=vbody&"<mark><b>Sparkior Media- Web Enquiry : </b></mark><br>"
				vbody=vbody&"<hr>"
				vbody=vbody&"<b>Name : </b>"&request("user_name")&"<br>"
				vbody=vbody&"<b>Phone : </b>"&request("user_email")&"<br>"
				vbody=vbody&"<b>Email : </b>"&request("user_phone") & "<br>"
				vbody=vbody&"<b>Message : </b>"&request("user_message")&"<br>"
				vbody=vbody&"<hr>"
			
				Dim myMail
				Set myMail=CreateObject("CDO.Message")
				myMail.Subject = "Sparkior Media - Web Enquiry"
				myMail.From="arun@sparkiormedia.in"
				myMail.To="arun@sparkiormedia.in"		
				myMail.HTMLBody= vbody
				'myMail.AddAttachment "https://www.sparkiormedia.in/images/logo.png"
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.livemail.co.uk"
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=587
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") =1
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") ="arun@sparkiormedia.in"
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") ="Arun@123"
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
				myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				myMail.Configuration.Fields.Update
				myMail.Send
				set myMail=nothing
								
				response.write "<span class='btn btn-primary btn-md'>Mail Sent!</span>"
%>
