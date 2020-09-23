<div align="center">

## CDO Mailer Script


</div>

### Description

Kept getting an error with ASP that I couldnt debug. So I came across this script. also ported it to PHP aswell. Hope its useful to someone as it was me.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ï¿½e7eN](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/e7en.md)
**Level**          |Beginner
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__4-1.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/e7en-cdo-mailer-script__4-9325/archive/master.zip)





### Source Code

```
<%
Dim objCDOSYSCon
Set objCDOSYSMail = Server.CreateObject("CDO.Message")
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
with objCDOSYSCon
	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	.Fields.Update
end with
with objCDOSYSMail
	Set .Configuration = objCDOSYSCon
	.From = "email@address.com"
	.To = "email@address.com"
	.Subject = "Insert Subject"
	.HTMLBody = "Insert Body"
	.Send
end with
Response.Write "Your email has been sent."
Set objCDOSYSMail = Nothing
Set objCDOSYSCon = Nothing
%>
```

