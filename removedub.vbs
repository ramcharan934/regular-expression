option explicit
dim a,newa,item,element,n,flag,str
a=array(5,6,7,7,6,8)
newa=array()
redim preserve newa(ubound(newa)+1)
newa(ubound(newa))=a(0)
for each item in a
	flag=true
	for each element in newa
		if(item = element)then
			flag=false
			exit for
		end if
	next
	if(flag)then
		redim preserve newa(ubound(newa)+1)
		newa(ubound(newa))=item
	end if
next
for each n in a
	str=str&" "&n
next

str=str&vbcr

for each n in newa
	str=str&" "&n
next
msgbox str


