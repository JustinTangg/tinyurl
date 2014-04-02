class tinyurl:
	import xlwt
	import xlrd
	import os
	import sys
	from urllib import urlopen
	excelloc = raw_input("Excel File Location Address: ")
	wb = xlrd.open_workbook(excelloc)
	sh = wb.sheet_by_index(0)
	wbk = xlwt.Workbook()
	sheet = wbk.add_sheet('TinyURL Sheet', cell_overwrite_ok = True)
	print "Working..."
	count = 0
	for rownum in range(sh.nrows):
		count += 1
		if sh.row_values(rownum)[0][:4] == "http":
			url = "http://tinyurl.com/api-create.php?url=" + sh.row_values(rownum)[0]
		else:
			url = "http://tinyurl.com/api-create.php?url=http://" + sh.row_values(rownum)[0]
		data = urlopen(url)
		sheet.write(rownum,0,sh.row_values(rownum)[0])
		data1 = data.readlines()[0].decode()
		errorcount = 0
		while data1 == "error" and errorcount < 5:
			time.sleep(1)
			errorcount += 1
			if sh.row_values(rownum)[0][:4] == "http":
				url = "http://tinyurl.com/api-create.php?url=" + sh.row_values(rownum)[0]
			else:
				url = "http://tinyurl.com/api-create.php?url=http://" + sh.row_values(rownum)[0]
			data = urlopen(url)
			data1 = data.readlines()[0].decode()
		sheet.write(rownum,1,data1)
		sys.stdout.write('\r' + str('%.1f'%((float(count)/float(sh.nrows))*100)) + "%")
		sys.stdout.flush()
	wbk.save(os.path.dirname(excelloc) + '\TinyURL.xls')
	print
	print "TinyURLs Creation Completed. Check " + os.path.dirname(excelloc) + " for TinyURL.xls"
	os.system('pause')