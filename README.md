<div align="center">

## Dynamic Form Field Capture


</div>

### Description

This code will caputre all fields in an HTML form and create a template for a VBScript function that passes the fields to a stored procedure.
 
### More Info
 
To use this file: Set the action on your form to point to this file.

You will then get a formatted list to process your form in VBScript and

a matching SQL Stored Procedure field list. Ex: action="get_fields.asp" target=_blank <br>

Warning: You should have your form fields named the same as your database

fields or this script won't really work for you... Have Fun!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Scott King](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/scott-king.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |HTML, VbScript \(browser/client side\)

**Category**       |[Controls/ Forms/ Dialogs/ Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/controls-forms-dialogs-menus__4-3.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/scott-king-dynamic-form-field-capture__4-8346/archive/master.zip)





### Source Code

```
<%@Language="VBScript"%>
<!--
	Created By: J. Scott King 4/28/2003
	Purpose: To make it easier for one to get form fields from an HTML form and turn
	them into scripting objects. Thus making a "plug and play" type interface for
	creating SQL stored procedures.
	Contact: jscott38@cox.net
-->
<title>Dynamic Form Field Creation</title>
To use this file: Set the action on your form to point to this file.
You will then get a formatted list to process your form in VBScript and
a matching SQL Stored Procedure field list. Ex: action="get_fields.asp" target=_blank <br>
Warning: You should have your form fields named the same as your database
fields or this script won't really work for you... Have Fun!
<hr>
<table width=100% cellspacing=1 bgcolor=Gainsboro>
<tr>
 <td colspan=2>
	<b>VBScript Form Fields.....</b>
	<textarea name=vbform rows=10 cols=110>
	<%
	For i = 1 to Request.Form.Count
	  fieldName = Request.Form.Key(i)
	  fieldValue = Request.Form.Item(i)
	  If i Mod 3 = 0 then
			Response.Write(" Request.Form(&quot;"& fieldName& "&quot;) &amp; &quot;','&quot; &amp; _ "&vbCrLf)
	  Else
			Response.Write(" Request.Form(&quot;"& fieldName& "&quot;) &amp; &quot;','&quot; &amp;")
		End If
	Next
	%>
	</textarea>
 </td>
</tr>
<tr>
 <td>
	<b>SQL Stored Procedure Fields.....</b>
	<textarea name=spfields rows=20 cols=40>
	<%
	For ix = 1 to Request.Form.Count
	  fieldName = Request.Form.Key(ix)
	  fieldValue = Request.Form.Item(ix)
	  Response.Write("@"&fieldName&" varchar(50),"&vbCrLf)
	Next
	%>
	</textarea>
 </td>
 <td>
	<b>SQL Stored Procedure Update Fields.....</b>
	<textarea name=spupdate rows=20 cols=40>
		<%
		For ix = 1 to Request.Form.Count
		  fieldName = Request.Form.Key(ix)
		  fieldValue = Request.Form.Item(ix)
		  Response.Write(fieldname & " = @"&fieldName&","&vbCrLf)
		Next
		%>
	</textarea>
 </td>
</tr>
</table>
```

