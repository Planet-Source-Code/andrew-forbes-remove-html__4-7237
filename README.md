<div align="center">

## Remove HTML


</div>

### Description

The trimHTML(string) function removes any html tags that are embedded in the string.
 
### More Info
 
a string

Useage:

x = trimHTML("a string with <b>html</b>")

which produces: "a string with html"

the string, with all html elements removed.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andrew Forbes](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andrew-forbes.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andrew-forbes-remove-html__4-7237/archive/master.zip)





### Source Code

```
<%
function TrimHTML(strHTML)
	dim iteration
	iteration=0
	do until instr(1,strHTML,"<")=0 and instr(1,strHTML,">")=0
		b=instr(1,strHTML,"<")
		if b>0 then
			c = instr(b+1,strHTML,">")
			if c>0 then
				retVal = mid(strHTML,b,c-(b-1))
				strHTML=Replace(strHTML,retVal,"")
			end if
		end if
		iteration = iteration + 1
		if iteration=1000 then exit do
	loop
	TrimHTML= strHTML
End function
dim str
str="952-91</font><font size="+chr(34)+"+1"+chr(34)+" color="+chr(34)+"#0000FF"+chr(34)+">7-0489</font>"
%>
Original text with html: <% =str %><br>
text with html removed: <% =TrimHTML(str) %>
```

