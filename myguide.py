import mammoth
from bs4 import BeautifulSoup


def convert2html(file):
	with open(file,'rb') as docx_file:
		result = mammoth.convert_to_html(docx_file)
		html = result.value #generated html
		messages = result.messages #any messages

	soup = BeautifulSoup(html,'html.parser')
	

	info_head = """<div style="display: none;">
	#EMCARDS_CSS
	#EXPAND_ELEMENT_JS
	#EXPANDABLE_JS
	#PLUS_MINUS_CURSOR
</div>
<div class="row">
	<div class="col-md-12">"""

	title_head = """<div class="panel panel-default">
			<div class="panel-heading" style="font-size: 1.2em;">
				<span class="glyphicon glyphicon-minus" style="padding-left: 5px;"></span>
				</div>
				"""

	title_tail = """"""

	body_head = """<ul class="list-group" style="display: block;">
				<li class="list-group-item">
				"""

	body_tail = """</li>
			</ul></div>"""


	html = html.replace('<p>#h#</p>','<rhead>')
	html = html.replace('<p>#/h#</p>','</rhead>')

	html = html.replace('<p>#t#</p>','<rtitle>')
	html = html.replace('<p>#/t#</p>','</rtitle>')

	html = html.replace('<p>#b#</p>','<rbody>')
	html = html.replace('<p>#/b#</p>','</rbody>')
	
	if '#t#' in html:
		html = html.replace('<br />#t#<br />','<rtitle><p>')
		html = html.replace('<br />#t#','<rtitle><p>')
		html = html.replace('<p>#t#<br />','<rtitle><p>')
		html = html.replace('<p>#t#','<rtitle><p>')
	if '#/t#' in html:
		html = html.replace('<br />#/t#<br />','</p></rtitle>')
		html = html.replace('<br />#/t#','</p></rtitle>')
		html = html.replace('#/t#<br />','</p></rtitle>')
		html = html.replace('#/t#</p>','</p></rtitle>')
	if '#b#' in html:
		html = html.replace('<p>#b#<br />','<rbody><p>')
		html = html.replace('#b#<br />','<rbody><p>')
		html = html.replace('<br />#b#','<rbody><p>')
		html = html.replace('#b#','<rbody><p>')
	if '#/b#' in html:
		html = html.replace('<br />#/b#',',</p></rbody>')
		html = html.replace('<p>#/b#</p>',',</rbody>')
		
	soup = BeautifulSoup(html,'html.parser')

	title = soup.rhead.text
	html = html.replace(title,'')

	soup = BeautifulSoup(html,'html.parser')

	top = """<div style="display: none;">
	#EMCARDS_CSS
	#EXPAND_ELEMENT_JS
	#EXPANDABLE_JS
	#PLUS_MINUS_CURSOR
</div>
<h2 style="text-align: center;">[Title]</h2>
<div class="row">
	<div class="col-md-12">"""

	top = top.replace('[Title]',title)

	bottom = """
		</div>"""

	top_start_pannel = """<div class="panel panel-default">
	<div class="panel-heading" style="font-size: 1.2em;">
		<span class="glyphicon glyphicon-minus" style="padding-left: 5px;"></span>
		<b>"""

	top_end_pannel = """		</b>
	</div>"""

	bottom_start_pannel = """	<ul class="list-group" id="summaryList" style="display: block;">
		<li class="list-group-item">
			<span>"""

	bottom_end_pannel = """			</span>
		</li>
	</ul>
</div>"""

	html_body = ''

	soup = BeautifulSoup(top+str(soup),'html.parser')
	remove_list = ['<p>','<strong>','</strong>','</p>','<br />','<br/>']

	for i in soup.find_all('rtitle'):
		souptext = str(i)
		for j in remove_list:
			souptext = souptext.replace(j,'')
		soup = BeautifulSoup(str(soup).replace(str(i),top_start_pannel+souptext+top_end_pannel),'html.parser')


	for i in soup.find_all('rbody'):
		soup = BeautifulSoup(str(soup).replace(str(i),bottom_start_pannel+str(i)+bottom_end_pannel),'html.parser')

	print(soup.prettify())

convert2html('S:\Palliative.docx')
# convert2html('S:\sepsis.docx')
