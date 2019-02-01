import os
import xlrd
import base64
from bs4 import BeautifulSoup
from collections import OrderedDict
from PIL import Image, ImageDraw, ImageFont

''' Script path and relative paths to SVG and Excel sheet '''
script_dir = os.path.dirname(__file__)

''' Path for personalized SVGs and PNGs '''
svg_path = r'renders/svg'
png_path = r'renders/png'

workbook = xlrd.open_workbook('contact_data/sheet/contacts.xlsx')
sheet1 = workbook.sheet_by_index(0)

''' Excel to dict '''
kids = OrderedDict()
for row_number in range(1, sheet1.nrows):
	contact_data = OrderedDict()
	row_values = sheet1.row_values(row_number)
	contact_data['email'] = row_values[2]
	contact_data['program'] = row_values[1]
	contact_data['image'] = row_values[3]
	kids[row_values[0]] = contact_data

''' SVG Manipulation '''
svg_template = open ("templates/svg/invitation.svg", 'r+').read()
soup = BeautifulSoup(svg_template, 'xml')
target = soup.find('flowPara')

''' SVG and PNG Rendering '''
for name, info in kids.items():
	# renders/svg/name.svg
	if not os.path.exists(svg_path):
		os.makedirs(svg_path)
	filepath = '{}/{}.svg'.format(svg_path, name)
	try:
		with open (filepath, 'w') as file:
			new_tag = soup.new_tag("text")
			new_tag.string = "Get a chance to witness {}'s accomplishments in the {}!".format(name, info["program"])
			target.clear()
			target.insert(1, new_tag)
			file.write(str(soup))
		print("{}.svg - OK.".format(name))
	except Exception as e:
		print("{}.svg - FAIL.".format(name))
		raise e

	# renders/png/name.png
	if not os.path.exists(png_path):
		os.makedirs(png_path)
	filepath = '{}/{}.png'.format(png_path, name)
	try:
		image = Image.open('templates/png/invitation.png')
		draw = ImageDraw.Draw(image)
		image_width, image_height = image.size

		def draw_text(line, h_offset=0, v_offset=0, caps=False, 
			align=None, colour = '#ffffff', font_size=45):
			font = ImageFont.truetype('fonts/Rubik/Rubik-Bold.ttf', size=font_size)
			if caps:
				line = line.upper()
			text_width, text_height = draw.textsize(line, font=font)
			if align == "centre":
				text_position = (((image_width-text_width)/2)+h_offset, 
					((image_height-text_height)/2)+v_offset)
			else:
				text_position = ((image_width-40+h_offset), (image_height-40+v_offset))
			draw.text(text_position, line, fill=colour, font=font)

		draw_text("Hey! "+name+" here!", v_offset=-1200, caps=True, align='centre', colour='#f1f442', font_size=145)
		draw_text("Come check out", v_offset=-1000, caps=True, align='centre', font_size=125)
		draw_text("what I've done at", v_offset=-800, caps=True, align='centre', font_size=125)
		draw_text(info["program"], v_offset=-600, caps=True, align='centre', colour='#f1f442', font_size=145)

		# Grabs the image using the file name from the sheet and resizes it to fit
		try:
			kid_image = Image.open('contact_data/pics/{}.png'.format(info["image"]))
		except:
			kid_image = Image.open('contact_data/pics/{}.jpg'.format(info["image"]))
		kid_image.thumbnail((1920, 1080), Image.ANTIALIAS)

		# Pastes the kid's image onto the generic background
		image.paste(kid_image, (int((image_width-kid_image.size[0])/2), image_height-1900))
		image.save(filepath)

		print("{}.png - OK.".format(name))
	except Exception as e:
		print("{}.png - FAIL.".format(name))
		raise e
	finally:
		print ("\n")
		text = None