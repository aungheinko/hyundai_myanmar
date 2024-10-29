import os
from tkinter import Tk
from tkinter import filedialog
from PIL import Image, ImageDraw, ImageFont

# Function to ask user to select a directory
def ask_directory(title):
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_selected = filedialog.askdirectory(title=title)
    return folder_selected

# Function to add a text watermark
def add_text_watermark(image, text, position, font_size=30):
    # Create an editable image (RGBA to support transparency)
    watermark_image = image.convert("RGBA")
    
    # Make the watermark overlay
    txt_layer = Image.new("RGBA", watermark_image.size, (255,255,255,0))
    draw = ImageDraw.Draw(txt_layer)
    
    # Set up font (you may need to specify the full path to the font file)
    try:
        font = ImageFont.truetype("arial.ttf", font_size)  # Use your desired font and size
    except IOError:
        font = ImageFont.load_default()
    
    # Add text to the image
    draw.text(position, text, font=font, fill=(255, 255, 255, 128))  # White text with some transparency
    
    # Combine the watermark with the image
    watermarked_image = Image.alpha_composite(watermark_image, txt_layer)
    
    # Convert back to RGB to save as JPEG
    return watermarked_image.convert("RGB")

# Get the input directory (where the .jfif files are)
input_directory = ask_directory("Select the folder containing .jfif files")

# Get the output directory (where the .jpeg files should be saved)
output_directory = ask_directory("Select the folder where .jpeg files will be saved")

# Create the output directory if it doesn't exist
if not os.path.exists(output_directory):
    os.makedirs(output_directory)

# Loop through all files in the input directory
for filename in os.listdir(input_directory):
    # Check if the file is a .jfif file
    if filename.endswith('.jfif'):
        # Open the image file
        with Image.open(os.path.join(input_directory, filename)) as img:
            # Convert the image to .jpeg
            new_filename = filename.replace('.jfif', '.jpeg')
            
            # Add watermark (customize the text and position)
            watermarked_img = add_text_watermark(img, text="Style Inspiration Daily", position=(30, 30))
            
            # Save the image in .jpeg format
            watermarked_img.save(os.path.join(output_directory, new_filename), 'JPEG')

        print(f"Converted {filename} to {new_filename} with watermark")

print("All conversions with watermarks are complete!")
