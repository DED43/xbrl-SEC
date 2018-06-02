import os
from xml.dom import minidom
from urllib.request import urlopen
from bs4 import BeautifulSoup
from pprint import pprint
import pandas as pd
from datetime import datetime

class Company:
	def __init__(self, cik):
		self.cik = cik
		self.documents = []
		self.data = {}
		url_string = "http://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK="+self.cik+"&type=10-%25&dateb=&owner=exclude&start=0&count=400&output=atom"
        
		self.forms = [
            '10-K', '10-K/A', '10-KT', '10-KT/A',
            '10-Q', '10-Q/A', '10-QT', '10-QT/A']
		self.datapath = os.path.join( os.path.realpath('.'), '/'.join(['data', cik]))
		if not os.path.exists( self.datapath):
			os.makedirs( self.datapath)

		xml_file = urlopen(url_string)
		self.xml = minidom.parse(xml_file)

	def getDocuments(self):
		"""
		Crawls edgar to get a list of all 10-Q/K XBRL files
		"""
		xml = self.xml

		mapping = {	'name':'conformed-name',
					'fiscal_year':'fiscal-year-end',
					'state_location':'state-location',
					'state_incorporation':'state-of-incorporation',
					'sic':'assigned-sic',
					'sic_desc':'assigned-sic-desc',
					'cik':'cik' }

		for k, v in mapping.items():
			try:
				self.data[k] = xml.getElementsByTagName(v)[0].firstChild.nodeValue
			except:
				pass

		listDocs = xml.getElementsByTagName('entry')
		for document in listDocs: 
			try: 
				parent = document.getElementsByTagName('content')[0]
				if parent.getElementsByTagName('filing-type')[0].firstChild.nodeValue in self.forms:
					parent.getElementsByTagName('xbrl_href')[0].firstChild.nodeValue
					doc = Document(parent)
					doc.data['file'] = self.datapath + '/{}_{}.{}'.format(
						str(doc.data['filing_type']).replace('/',''), 
						doc.data['filing_date'],
						doc.data['xbrl_url'].strip().split('.')[-1])
					try:
						f = open( doc.data['file'], 'x+b')
						f.write( urlopen( doc.data['xbrl_url']).read())
						f.close()
					except FileExistsError: #file already exists
						pass
					self.documents.append(doc)
			except IndexError: #no xbrl data
				pass

	def getSeries(self, codes):
		'''
		'''
		if len(self.documents) == 0:
			self.getDocuments()
		if len(self.documents) == 0:
			raise Exception("No data available from Edgar")
		documents = self.documents

		collected = {}
		for code in codes:
			collected[code] = {}
		result = collected
		
		for document in documents:
			datas = document.getItems(codes)
			for code in datas:
				for period in datas[code]:
					if period in collected[code]:
						if collected[code][period][0] != datas[code][period]:
							print('Code: ', code,'. Period: ', period,'. Err: different values',
							' x: ', collected[code][period][0], '(',collected[code][period][1],')',
							' xx: ', datas[code][period], '(',document.data['fixing_date'],')')
							if document.getDate(document.data['fixing_date']) < document.getDate(collected[code][period][1]):
								collected[code][period] = [datas[code][period],document.data['fixing_date']]
					else:
						collected[code][period] = [datas[code][period],document.data['fixing_date']]

		for code in collected:
			for period in collected[code]:
				result[code][period] = collected[code][period][0]


		return pd.DataFrame(result)



class Document:
	'''
	'''
	def __init__(self, filing):
		'''
		'''
		self.data = {}
		self.data['date_format'] = '%Y-%m-%d'
		self.data['filing_date'] = filing.getElementsByTagName('filing-date')[0].firstChild.nodeValue
		self.data['filing_type'] = filing.getElementsByTagName('filing-type')[0].firstChild.nodeValue
		self.data['filing_url'] = filing.getElementsByTagName('filing-href')[0].firstChild.nodeValue
		self.data['xbrl_url'] = self.getXBRLurl()
		self.data['file'] = None
		self.data['fixing_date'] = None

	def getXBRLurl(self):
		filing = urlopen(self.data['filing_url']).read()
		soup = BeautifulSoup(filing,'lxml')
		xbrl_table = soup.findAll('table', attrs={'summary':"Data Files"})[0]
		return 'http://www.sec.gov'+xbrl_table.findAll('a')[0]['href']

	def getDate(self, text):
		try:
			date = datetime.strptime( text, self.data['date_format'])		
		except ValueError:
			print('not date:', text)
		return date

	def getPeriod(self, end: str, start: str = ''):
		period = ''
		if start == '':
			date = self.getDate( end)
			if date.month in [3,6,9]:
				period = '{}Q{}'.format( date.year, int(date.month/3))
			elif date.month == 12:
				period = '{}Y'.format( date.year)
		else:
			date0 = self.getDate(start)
			date1 = self.getDate(end)
			if date1.month - date0.month == 11:
				period = '{}Y'.format( date1.year)
			elif date1.month - date0.month == 2:
				period = '{}Q{}'.format( date1.year, int(date1.month/3))
			elif date1.month - date0.month == 5:
				period = '{}H{}'.format( date1.year, int(date1.month/6))
			elif date1.month - date0.month == 8:
				period = '{}m9'.format( date1.year)
		return period

	def getContext(self, soup, contextRef):
		contexts = soup.getElementsByTagNameNS('*','context')
		data = {}
		for context in contexts:
			if context.attributes['id'].value == contextRef:
				try: #instant
					date = context.getElementsByTagNameNS('*','instant')[0].firstChild.nodeValue
					data['end'] = date
					data['period'] = self.getPeriod( date)
				except: #period
					start = context.getElementsByTagNameNS('*','startDate')[0].firstChild.nodeValue
					end = context.getElementsByTagNameNS('*','endDate')[0].firstChild.nodeValue
					data['start'] = start
					data['end'] = end
					data['period'] = self.getPeriod( end, start)
				
				try:
					data['segment'] = context.getElementsByTagNameNS('*','explicitMember')[0].firstChild.nodeValue
				except:
					data['segment'] = 'root'
					pass
		return data

	def getItems(self, codes): #get consolidated items
		xbrl_data = open( self.data['file'], 'r')
		soup = minidom.parse(xbrl_data)

		_contexts = soup.getElementsByTagNameNS('*','context')
		contexts = {}
		for context in _contexts:
			data = {}
			try: #instant
				date = context.getElementsByTagNameNS('*','instant')[0].firstChild.nodeValue
				data['end'] = date
				data['period'] = self.getPeriod( date)
			except: #period
				start = context.getElementsByTagNameNS('*','startDate')[0].firstChild.nodeValue
				end = context.getElementsByTagNameNS('*','endDate')[0].firstChild.nodeValue
				data['start'] = start
				data['end'] = end
				data['period'] = self.getPeriod( end, start)
			
			try: #check whether this context is for segment data
				context.getElementsByTagNameNS('*','explicitMember')[0].firstChild.nodeValue
			except: #not segment data context
				contexts[context.attributes['id'].value] = data
				pass

		datapoints = {}
		for code in codes:
			datapoints[code] = {}
			code_soup = soup.getElementsByTagNameNS('*', code)
			for point in code_soup:
				contextRef = point.attributes['contextRef'].value
				if contextRef in contexts:
					period = contexts[contextRef]['period']
					datapoints[code][period] = point.firstChild.nodeValue
					end_date = contexts[contextRef]['end']
					if self.data['fixing_date'] is None or self.getDate(end_date) > self.getDate(self.data['fixing_date']):
						self.data['fixing_date'] = end_date
		return datapoints


#db = pd.ExcelFile('data.xlsx')
dat = pd.read_excel( 'in.xlsx', sheet_name='descr')
writer = pd.ExcelWriter('data.xlsx')
for cik in dat.transpose().index:
	print( cik)
	codes = []
	for code in dat.index:
		if dat.at[code,cik]==1: codes.append(code)
	comp = Company(cik)
	data = comp.getSeries(codes)
	data.to_excel( writer, cik, freeze_panes = (1,1))
writer.save()


