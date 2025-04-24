from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import pandas as pd
import os

# Use a Korean TTF font (macOS path)
font_path = "/System/Library/Fonts/Supplemental/AppleGothic.ttf"
pdfmetrics.registerFont(TTFont('AppleGothic', font_path))

# read the data source for each label
df = pd.read_csv('905604.csv')
addresses = df[['Name', 'Congregation']]

# Create a new canvas
c = canvas.Canvas("result.pdf", pagesize=letter)

page_width, page_height = letter # 612.0,792.0
     	
# 'height' and 'width' conversions:
# 1 mm = 2.835 pts
# 1 cm = 28.346 pts
# 1 inch = 72 pts
# 0.5 inch = 36 pts
# 3 mm = 0.3 cm = 8.505 pts
# 5 mm = 0.5 cm = 14.175 pts

num_columns = 2
num_rows = 5
column_width = 229.6026
row_height = 144.5646
left_margin = 0
left_text_margin = 42.519
top_margin = 0
bottom_margin = 34.29866
top_text_margin = 105 #107.7148
right_label_gap = 0 #28.346

total_items = len(addresses)
items_per_page = num_columns * num_rows
pages = (total_items + items_per_page - 1) // items_per_page  # Ceiling division

# draw horizontal lines
for row in range(pages):
    c.setLineWidth(0.1)
    c.line(0,page_height - top_margin - (row * row_height),page_width,page_height - top_margin - (row * row_height))

# draw vertical lines
for v in range(num_columns*2):
    x_vertical_position = v*column_width + (left_margin*v) + (right_label_gap*v)
    c.line(x_vertical_position,0,x_vertical_position,page_height)

# Loop through each address, positioning them in a 'num_columns x num_rows' grid
for index, address in addresses.iterrows():
    column = index % num_columns  # Determine the current column
    row = index // num_columns % num_rows  # Determine the current row
 
    # Calculate x and y position for the current label
    x_position = (column * column_width) + left_margin + left_text_margin + (right_label_gap * column)
    y_position =  page_height - top_margin - (row * row_height) - (top_text_margin)
    	
    # Start the text object at the calculated position
    text = c.beginText(x_position, y_position)
    text.setFont("AppleGothic", 14)  # Set the font and size
    text.textLine(address['Name']) 
    c.drawText(text)

    text = c.beginText(x_position, y_position-22)
    text.setFont("AppleGothic", 12)  # Set the font and size
    text.textLine(address['Congregation']) 
    c.drawText(text)

    # Move to the next page if necessary
    if (index + 1) % (num_columns * num_rows) == 0:
        c.showPage()
        # draw horizontal lines
        for row in range(pages):
            c.setLineWidth(0.1)
            c.line(0,page_height - top_margin - (row * row_height),page_width,page_height - top_margin - (row * row_height))
        
        # draw vertical lines
        for v in range(num_columns*2):
            x_vertical_position = v*column_width + (left_margin*v) + (right_label_gap*v)
            c.line(x_vertical_position,0,x_vertical_position,page_height)
        
# Save the PDF
c.save()
