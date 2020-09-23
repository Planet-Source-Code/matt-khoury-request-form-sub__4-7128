<div align="center">

## Request Form Sub


</div>

### Description

This code will help make your code more common, and provide simpler viewing for the Request.Form collection, by allowing the developer to place the entire Request.Form collection into a single Sub, therefore helping to organize your code a little better. We’ve all seen the type of code where Request.Form’s are done many, many time throughout an ASP. It’s hard to keep track of, right? Well, using this Sub will help you to organize your code, and your form collection.
 
### More Info
 
Request.Form Collection

Request.Form Variables

None that I am aware of.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Matt Khoury](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/matt-khoury.md)
**Level**          |Intermediate
**User Rating**    |3.3 (13 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[System Services/ Functions](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/system-services-functions__4-23.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/matt-khoury-request-form-sub__4-7128/archive/master.zip)





### Source Code

```
<%
Dim strExample1, strExample2, strExample3
Call GetPostedItems(strExample1, strExample2, strExample3)
%>
If you want, here is where you can Response.Write
the variables to view the items in your collection:<br>
<%
	Response.Write "strExample1 = [" & strExample1 & "]<br>"
	Response.Write "strExample2 = [" & strExample2 & "]<br>"
	Response.Write "strExample3 = [" & strExample3 & "]<br>"
Sub GetPostedItems(ByRef strExample1, ByRef strExample2, _
						ByRef strExample3)
	'-- Enter the NAME of the field below.
	strExample1 = Request.Form("Example1")
	strExample2 = Request.Form("Example2")
	strExample3 = Request.Form("Example3")
End Sub
%>
```

