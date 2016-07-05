str=inputbox("enter a sentence")
arr=split(str," ")


for each item in arr	
	tempstr=ucase(left(item,1))&mid(item,2)
	temp=temp&" "&tempstr
next
msgbox temp