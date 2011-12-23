#!/usr/bin/env python
#
#   gen-cards.py -- Typeset Christmas card envelopes using a list sourced
#					from a Google Docs spreadsheet
#
#  <andy@payne.org>
#  Dec-2009

import gdata.docs
import gdata.docs.service
import gdata.spreadsheet.service
import re, os, os.path
import ConfigParser
import sys

# Read configuration
cfile = os.path.expanduser('~/.christmascard')
if not os.path.isfile(cfile):
	print "Configuration file %s is missing!" % cfile
	sys.exit(1)

c = ConfigParser.ConfigParser()
c.read(cfile)
config = {}
for i in ('username', 'password', 'doc_name', 'addr1', 'addr2', 'addr3'):
	config[i] = c.get('general', i).strip()

# Connect to Google
gd_client = gdata.spreadsheet.service.SpreadsheetsService()
gd_client.email = config['username']
gd_client.password = config['password']
gd_client.source = 'payne.org-ccard-1'
gd_client.ProgrammaticLogin()

# Query for the rows
print "Reading rows...."
q = gdata.spreadsheet.service.DocumentQuery()
q['title'] = config['doc_name']
q['title-exact'] = 'true'
feed = gd_client.GetSpreadsheetsFeed(query=q)
assert(feed)
spreadsheet_id = feed.entry[0].id.text.rsplit('/',1)[1]
feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
worksheet_id = feed.entry[0].id.text.rsplit('/',1)[1]
rows = gd_client.GetListFeed(spreadsheet_id, worksheet_id).entry

# -------------------------------------------------------------------------
#
#  Generate the LaTeX document
#

fd = open("document.tex", "w")

fd.write(r"""
\documentclass[letter,landscape]{article}
\usepackage[papersize={7.5in,4.75in},margin=0.25in]{geometry}
\usepackage[pdftex]{graphicx}
\usepackage{cmbright}
%\usepackage[letter,center,frame,landscape]{crop}

\newcommand{\address}[1]{
    \setlength{\parindent}{0pt}
    \bigskip\bigskip
    \hbox{
        \raisebox{-3pt}{\includegraphics[width=0.45in,keepaspectratio]{ChristmasTree2.png}}
        \vbox{
""")

fd.write(r"""%s \\
%s \\ 
%s """ % (config['addr1'], config['addr2'], config['addr3']))

fd.write(r"""}
        \hfil
    }

    \vspace*{1.75in}
    \hbox{\hskip 3.0in \vbox{\Large #1}}
    \newpage
}

\pagestyle{empty}
\begin{document}
""")

dups = {}

for (count, row) in enumerate(rows):
	data = {}
	for key in row.custom:
		if row.custom[key].text:
			data[key] = row.custom[key].text.strip()
		else:
			data[key] = ""

	if len(data['address']) > 0 and str.lower(data['card']) == 'yes':
		data['zip'] = data['zip'].strip("'")
		t = [data['name'], data['address'], data['address2'], data['address3'], "%s, %s \\ %s" % (data['city'], data['state'], data['zip']), data['country']]
		#  Join and escape strings for LaTeX
		txt = """\\\\""".join([x for x in t if len(x) > 0])
		txt = re.sub("#","\\#",txt)
		txt = re.sub("&","\\&",txt)
		
		fd.write("""\\address{%s}\n""" % txt)
		
		# Basic dup check based on finding the last name
		name = str.lower(data['name']).split()
		while name[-1] in ('family', '&', 'and'):
			name.pop()

		key = "%s-%s" % (name[-1], data['zip'])
		if key in dups:
			print "***** POSSIBLE DUP"
			print "     ", count, " ".join(t)
			print "     ", dups[key]
			print
			print
			
		dups[key] = "%d %s" % (count, " ".join(t))

fd.write("\end{document}\n")    
fd.close()

# Run LaTeX to make the PDF to print
os.system("pdflatex document.tex")

