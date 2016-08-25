'''
Crawl web server file directory and download all to the local

use: crawlDir.py <url> <dest>
'''

import requests,os,sys
from HTMLParser import HTMLParser

class crawlHtml(HTMLParser):

	def __init__(self):
		HTMLParser.__init__(self)
		self.type = ''
		self.path = []

	def handle_starttag(self, tag, attrs):
		if tag == 'img':
			for (key,value) in attrs:
				if key == 'alt':
					if value not in ['[ICO]','[PARENTDIR]']:
						self.type = 'dir' if value == '[DIR]' else 'file'
		if tag == 'a' and self.type:
			for (key,value) in attrs:
				if key == 'href':
					self.path.append([self.type,value])

def slash(url):
	return url if url[-1]=='/' else url+'/'

def getFiles(url,path):
	url = slash(url)
	path = slash(path)
	if not os.path.exists(path):
		os.mkdir(path)
	c = crawlHtml()
	c.feed(requests.get(url).content)
	for files in c.path:
		if files[0] == 'dir':
			getFiles(url+files[1],path+files[1])
		else:
			print 'downloading {0} => {1}'.format(url+files[1],path+files[1])
			with open(path+files[1],'wb') as f:
				f.write(requests.get(url+files[1]).content)

if __name__ == '__main__':
	# url = 'http://218.76.35.75:20104/.git'
	# path = '.git'
	# getFiles(url,path)
	getFiles(sys.argv[1],sys.argv[2])
	print '===>>>done !<<<==='