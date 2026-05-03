import numpy as np
from PIL import Image, ImageDraw

# Create a test image with 3 distinct blocks
img = Image.new('L', (300, 100), color=255)
draw = ImageDraw.Draw(img)
draw.rectangle([10, 10, 80, 90], fill=0)
draw.rectangle([110, 10, 180, 90], fill=0)
draw.rectangle([210, 10, 280, 90], fill=0)

arr = np.array(img)
# 1 for content, 0 for white
content = (arr < 250).astype(int)

# Column projection
col_sum = content.sum(axis=0)
# Find boundaries
is_content = col_sum > 0
diff = np.diff(is_content.astype(int))
starts = np.where(diff == 1)[0] + 1
if is_content[0]: starts = np.insert(starts, 0, 0)
ends = np.where(diff == -1)[0] + 1
if is_content[-1]: ends = np.append(ends, len(col_sum))

print("Starts:", starts)
print("Ends:", ends)
