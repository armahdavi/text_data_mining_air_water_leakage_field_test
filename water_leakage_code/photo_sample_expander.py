# -*- coding: utf-8 -*-

"""
Program to replicate copies of a photo prototype (good for architectural drawing of simulation cases)
Created on Thu Jan 16 13:36:45 2025

@author: MahdaviAl
"""

from PIL import Image

def expand_image(image_path, L, W):
    # Open the prototype image
    prototype = Image.open(image_path)
    
    # Get the dimensions of the prototype image
    prototype_width, prototype_height = prototype.size
    
    # Calculate the dimensions of the expanded image
    expanded_width = prototype_width * W
    expanded_height = prototype_height * L
    
    # Create a new blank image with the calculated dimensions
    expanded_image = Image.new('RGB', (expanded_width, expanded_height))
    
    # Paste the prototype image into the expanded image L x W times
    for i in range(L):
        for j in range(W):
            expanded_image.paste(prototype, (j * prototype_width, i * prototype_height))
    
    return expanded_image

# Example usage
image_path = 'gypsum_sample.png'  # Path to the prototype image
L = 10  # Number of times to repeat the image vertically
W = 10 # Number of times to repeat the image horizontally

expanded_image = expand_image(image_path, L, W)
expanded_image.save(r'expanded_gypsum_sample.png')
expanded_image.show()
