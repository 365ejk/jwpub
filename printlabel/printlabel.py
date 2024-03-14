from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import pandas as pd

# read the data source for each label
df = pd.read_csv('list.csv')
addresses = df[['Name', 'Street', 'City', 'State', 'Zip']]

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

num_columns = 3
num_rows = 10
column_width = 189.9182
row_height = 72
left_margin = 14.175
left_text_margin = 8.505
top_margin = 36
bottom_margin = 36
top_text_margin = 8.505
right_label_gap = 8.505

# Loop through each address, positioning them in a 'num_columns x num_rows' grid
for index, address in addresses.iterrows():
    column = index % num_columns  # Determine the current column
    row = index // num_columns % num_rows  # Determine the current row
    
    # Calculate x and y position for the current label
    x_position = (column * column_width) + left_margin + left_text_margin + (right_label_gap * column)
    y_position =  page_height - top_margin - (row * row_height) - (top_text_margin*2)
    	
    # Start the text object at the calculated position
    text = c.beginText(x_position, y_position)
    text.setFont("Helvetica", 10)  # Set the font and size
    
    # Create the full address string and add it to the text object
    full_address = f"{address['Name']}\n{address['Street']}\n{address['City']}, {address['State']} {address['Zip']}"
    for line in full_address.split('\n'):
    	text.textLine(line)
        
    # Draw the text object onto the canvas
    c.drawText(text)

    # Move to the next page if necessary
    if (index + 1) % (num_columns * num_rows) == 0:
        c.showPage()

# Save the PDF
c.save()
