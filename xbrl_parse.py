import os
from xml.dom import minidom
from urllib.request import urlopen
from bs4 import BeautifulSoup
from pprint import pprint
import pandas as pd
from datetime import datetime
from win32com.client import Dispatch


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

		self.task = 1
		# Tasks:
		# 1 - parse new reports
		# 2 - parse all reports

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
					parent.getElementsByTagName('xbrl_href')[0].firstChild.nodeValue #check for XBRL data
					doc = Document(parent)
					doc.data['file'] = self.datapath + '/{}_{}.{}'.format(
						str(doc.data['filing_type']).replace('/',''), 
						doc.data['filing_date'],
						doc.data['xbrl_url'].strip().split('.')[-1])
					try:
						f = open( doc.data['file'], 'x+b')
						f.write( urlopen( doc.data['xbrl_url']).read())
						f.close()
						self.documents.append(doc)
						print('-  +', doc.data['file'])
					except FileExistsError: #file already exists
						if self.task == 2: 
							self.documents.append(doc)
						pass
			except IndexError: #no xbrl data
				pass

	def getSeries(self, codes, data):
		dataNew = {}
		if len(self.documents) == 0:
			self.getDocuments()
		if len(self.documents) == 0:
			print('--- No new data')
			return pd.DataFrame(dataNew)
		documents = self.documents

		collected = {}
		for code in codes:
			collected[code] = {}
			dataNew[code] = {}
		
		for document in documents:
			datas = document.getItems(codes)
			for code in datas:
				for period in datas[code]:
					isSet = True
					if period in collected[code]:
						isSet = False
						if collected[code][period][0] != datas[code][period]:
							#print('Code: ', code,'. Period: ', period,'. Err: different values',
							#' curr: ', collected[code][period][0], '(',collected[code][period][1],')',
							#'  new: ', datas[code][period], '(',document.data['fixing_date'],')')
							
							fixing_date_old = document.getDate(collected[code][period][1])
							fixing_date_new = document.getDate(document.data['fixing_date'])

							if fixing_date_old == fixing_date_new:
								filing_date_old = document.getDate(collected[code][period][2])
								filing_date_new = document.getDate(document.data['filing_date'])
								if filing_date_old < filing_date_new:
									isSet = True
							elif fixing_date_new < fixing_date_old:
								isSet = True
							
					if isSet:
						collected[code][period] = [
							datas[code][period],
							document.data['fixing_date'],
							document.data['filing_date'],
							document.data['period']]

		for code in collected:
			for period in collected[code]:
				try:
					value_old = data.at[period,code]
					value_new = collected[code][period][0]
					if value_old != value_new:
						if period == collected[code][period][3]:
							data.at[period,code] = value_new
				except KeyError:
					dataNew[code][period] = collected[code][period][0]
					pass

		return pd.DataFrame(dataNew)


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
		self.data['period'] = None

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

	def setPeriod(self):
		if self.data['fixing_date'] is not None:
			date = self.getDate( self.data['fixing_date'])
			if date.month in [3,6,9]:
				period = '{}Q{}'.format( date.year, int(date.month/3))
				if 'Q' not in self.data['filing_type']:
					print('*** No period', self.data['fixing_date'])
			elif date.month == 12:
				period = '{}Y'.format( date.year)
				if 'K' not in self.data['filing_type']:
					print('*** No period', self.data['fixing_date'])
		else:
			date = self.getDate( self.data['filing_date'])
			if 'K' in self.data['filing_type']:
				period = '{}Y'.format( date.year - 1)
			elif 'Q' in self.data['filing_type']:
				period = '{}Q{}'.format( date.year, int(date.month/3))
		self.data['period'] = period

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

		periods = soup.getElementsByTagNameNS('*','DocumentPeriodEndDate')
		if len(periods)>0:
			self.data['fixing_date'] = periods[0].firstChild.nodeValue
		else: print('*** no DocumentPeriodEndDate tag. File: ', self.data['file'])

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
		fixing_date = ''
		for code in codes:
			datapoints[code] = {}
			code_soup = soup.getElementsByTagNameNS('*', code)
			for point in code_soup:
				contextRef = point.attributes['contextRef'].value
				if contextRef in contexts:
					period = contexts[contextRef]['period']
					datapoints[code][period] = int(point.firstChild.nodeValue)

					if self.data['fixing_date'] is None:
						end_date = contexts[contextRef]['end']
						if fixing_date=='' or self.getDate(end_date) > self.getDate(fixing_date):
							fixing_date = end_date
		
		if self.data['fixing_date'] is None:
			self.data['fixing_date'] = fixing_date
		self.setPeriod()

		return datapoints



stateOutFile = 1
try:
	out = pd.ExcelFile('data.xlsx')
except FileNotFoundError: #file doesn't exist
	print('Output data file error. All files will be re-parsed')
	stateOutFile = 2
	pass

work = pd.read_excel( 'in.xlsx', sheet_name='descr')
writerResult = pd.ExcelWriter('data.xlsx')
for cik in work.columns:
	print( cik)
	comp = Company(cik)

	data = pd.DataFrame()
	if stateOutFile == 1 and cik in out.sheet_names:
		data = pd.read_excel( out, sheet_name=cik)
	else: 
		comp.task = 2
	
	codes = []
	for code in work.index:
		if work.at[code,cik] > 0: 
			codes.append(code)
			if work.at[code,cik] == 2:
				comp.task = 2
	
	dataNew = comp.getSeries(codes,data)
	data = pd.concat( [data, dataNew])
	data.sort_index()

	#delete empty columns
	data = data.dropna( axis='columns', how='all')
	for code in work.index:
		if work.at[code,cik] > 0: 
			if code in data.columns:
				work.at[code,cik] = 1
			else:
				work.at[code,cik] = 0

	data.to_excel( writerResult, cik, freeze_panes = (1,1))
writerResult.save()
writerResult.close()


#save results of work
xl = Dispatch("Excel.Application")
xl.Visible = True # otherwise excel is hidden
wb = xl.Workbooks.Open( os.path.realpath('.') + '\\' + 'in')
sh = wb.Worksheets('descr')
i_m = len(work.index)
j_m = len(work.columns)
for i in range(len(work.index)):
	for j in range(len(work.columns)):
		sh.Range('B2').Offset( 1 + i, 1 + j).Value2 = int(work.iat[ i, j])
wb.Save()
wb.Close()
xl.Quit()
