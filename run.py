import socket
import pandas as pd
import xlsxwriter
import dns.resolver
import dns.reversename

#Constants
excel_path = 'example.xlsx'
column_name = 'Domain'

# Instance of result
workbook = xlsxwriter.Workbook('result.xlsx')
worksheet = workbook.add_worksheet()

# Putting the headers
worksheet.write("A1", "DOMAIN")
worksheet.write("B1", "IP")
worksheet.write("C1", "DNS")

# Reading data from excel
data = pd.read_excel (excel_path)
rows = pd.DataFrame(data, columns= [column_name])

# Gambi :(
row = 1

for domain in rows['Domain']:
	HOST = domain
	ip = ''
	for qtype in ['NS']:
		try:
			ip = socket.gethostbyname(domain)			
			answers = dns.resolver.query(HOST, qtype, raise_on_no_answer=False)
			for answer in answers:
				result = answer
				row = row + 1	
				print(row)			
				worksheet.write("{}{}".format("A", row), domain)
				worksheet.write("{}{}".format("B", row), str(ip))	
				worksheet.write("{}{}".format("C", row), str(result))				
		except:
			result = "ERRO NO DOMÍNIO"
			ip = "ERRO NO DOMÍNIO"
			row = row + 1	
			print(row)				
			worksheet.write("{}{}".format("A", row), domain)
			worksheet.write("{}{}".format("B", row), str(ip))	
			worksheet.write("{}{}".format("C", row), str(result))
			continue
		
workbook.close()