from bs4 import BeautifulSoup as bs 
import requests
import xlrd

def save(url, value):
	wf.write(url+','+value+'\n')

def find_end_page(raw_url):
	initial=200
	final=700
	itteration=15
	
	count = 0
	try:
		while initial<final and count<itteration:
			res = int((initial+final)/2)
			url = "https://www.bing.com/search?q=site"+str('%3a')+raw_url+"&qs=n&sp=-1&pq=undefined&sc=0-23&sk=&cvid=73099500678F488A9A766C0267E93EB7&first="+str(res)+"&FORM=PERE"
			source = requests.get(url)
			soup = bs(source.content.decode(), 'lxml')
			soup1 = soup.find('div', {'id': 'b_tween'})
			#print('initial=', initial, '  final=', final, count)
			if soup1!=None:
				initial = res
				soup2 = soup1.find('span', {'class': 'sb_count'})
			elif soup1==None:
				final = res
			count = count+1
		save(raw_url, soup2.text)
	except Exception as e:
		print(e)


wf = open('Bing123.csv', 'a')	
workbook = xlrd.open_workbook('Domain-list-Sumit.xlsx')
worksheet = workbook.sheet_by_index(0)
rows = worksheet.nrows

for i in range(680, 685): 
	try:
		raw_url = worksheet.cell(i, 1).value
		url = url = "https://www.bing.com/search?q=site"+str('%3a')+raw_url+"&qs=n&sp=-1&pq=undefined&sc=0-23&sk=&cvid=73099500678F488A9A766C0267E93EB7&first="+str(1000)+"&FORM=PERE"
		source = requests.get(url)
		soup = bs(source.content.decode(), 'lxml')
		soup1 = soup.find('div', {'id': 'b_tween'})
		if soup1!=None:
			soup2 = soup1.find('span', {'class': 'sb_count'})
			print(i,'/',rows, 'a')
			save(raw_url, soup2.text)
		elif soup1==None:
			find_end_page(raw_url)
			print(i,'/',rows, 'b')
	except Exception as e:
		print(e)