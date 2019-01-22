import os
import xlrd
from cairosvg import svg2png
from bs4 import BeautifulSoup
from collections import OrderedDict

from wand.image import Image


''' Script path and relative paths to SVG and Excel sheet '''
script_dir = os.path.dirname(__file__)
svg =  open ("svg_template/drawing.svg", 'r+').read()

''' Path for personalized SVGs and PNGs '''
svgpath = r'renders/svg'
pngpath = r'renders/png'

workbook = xlrd.open_workbook('contact_sheet/contacts.xlsx')
sheet1 = workbook.sheet_by_index(0)

''' Excel to dict '''
kids = OrderedDict()

for row_number in range(1, sheet1.nrows):
	contact_data = OrderedDict()
	row_values = sheet1.row_values(row_number)
	contact_data['email'] = row_values[2]
	contact_data['program'] = row_values[1]
	kids[row_values[0]] = contact_data

''' SVG Manipulation '''
soup = BeautifulSoup(svg, 'xml')
text = soup.find('flowPara')

for name, info in kids.items():
	text.string = "Get a chance to witness {}'s accomplishments in the {}!".format(name, info["program"])
	
	# renders/svg/name.svg
	if not os.path.exists(svgpath):
		os.makedirs(svgpath)
	filepath = svgpath+'/{}.svg'.format(name)
	try:
		with open (filepath, 'w') as file:
			file.write(str(soup))
		print("{}.svg - OK.".format(name))
	except Exception as e:
		print("{}.svg - FAIL.".format(name))
		raise e

	# renders/png/name.png
	if not os.path.exists(pngpath):
		os.makedirs(pngpath)
	filepath = pngpath+'/{}.png'.format(name)
	svg_code = soup.find('svg')
	try:
		''' This one uses Cairo SVG '''
		#with open (filepath, 'w') as file:
		#	svg2png(bytestring = bytes(str(svg_code),'UTF-8'), 
		#		write_to=filepath)
		
		''' This one uses Wand - ImageMagick '''
		with Image(blob=bytes(str(svg_code),'UTF-8'), format="svg") as image:
			image.format = "png"
			image.save(filename=filepath)
		print("{}.png - OK.".format(name))
	except Exception as e:
		print("{}.png - FAIL.".format(name))
		raise e
	finally:
		print ("\n")

	# Email
	# TODO