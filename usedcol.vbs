dim excel,wb,sh,rc

set excel=createobject("excel.application")
msgbox typename(excel)
'excel.visible=true
'excel.workbooks.add()
'excel.activeworkbook.saveas("c:\my\names.xls")
'excel.visisble=false
set wb=excel.workbooks.open("c:\my\names.xls")
msgbox typename(wb)
set sh=wb.sheets("sheet1")
rc=sh.usedrange.rows.count
for i=1 to rc 
msgbox sh.cells(i,1).value
msgbox sh.cells(i,2).value
next
wb.save()
wb.close()