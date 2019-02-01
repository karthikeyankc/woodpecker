# Woodpecker 🐦
A python script to create personalized SVGs and PNGs based on SVG and PNG templates.

# Usage
- Clone the repo, use `pip3 install requirements.txt`, and run the script (python 3) with the `contacts.xls`, pictures, and templates (for SVG and PNG) in place.
- For any new templates, you might need to adjust the `draw_text` parametres based on your design. By default, texts are centre aligned.
- The script uses PIL's thumbnail method to resize the second image. The default size is `1920 × 1080`.

# Note
- In the `Image` column of the `contacts.xls` sheet, do not include the file extension.
- Allowed image formats - Jpeg and PNG.