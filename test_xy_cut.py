import numpy as np
from PIL import Image

def recursive_xy_cut(img_arr, min_spacing=10):
    boxes = []
    
    def _cut(x0, y0, x1, y1):
        # Base case check
        if x1 - x0 <= 10 or y1 - y0 <= 10:
            return
            
        sub_arr = img_arr[y0:y1, x0:x1]
        
        # Check column projection (vertical cuts)
        col_sum = sub_arr.sum(axis=0)
        is_empty = col_sum == 0
        
        # Find continuous empty gaps
        # We want to find gaps that are wider than min_spacing
        # To do this simply, let's just find any gap and split.
        # But wait, letters might have small gaps (like 'i', 't'). So min_spacing is needed.
        
        # Let's use a simpler heuristic: find connected components using scipy if we had it.
        pass

