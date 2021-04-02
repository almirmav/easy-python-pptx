# easy-python-pptx
Python utility for using python-pptx libraries to insert titles, texts of pandas dataframes to pptx slides.


### Example of using
```python
import easy_python_pptx as epp
from pptx import Presentation
from pptx.util import Cm, Pt, Inches
import pandas as pd
import numpy as np

# create presentation 16x9
prs = Presentation()
prs.slide_width = Inches(16)
prs.slide_height = Inches(9)

# adding slide with text
slide = epp.addSlideTitle(prs, 'Slide with Text')
epp.addTextToSlide(slide, 'Our text for slide', Inches(1), Inches(1), Inches(3), Inches(3))

# adding slide with list
epp.addSlideWithList(prs, 'Slide with List', 'Our list', ['item1', 'item2', 'item3'])

# adding slide with pandas dataframe
df1 = pd.DataFrame(np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9]]), columns=['a', 'b', 'c'])
slide = epp.addSlideTitle(prs, 'Slide with Table')
epp.addTableToSlide(slide, df1, Inches(1), Inches(1), Inches(3), Inches(1), colWidths=[3, 1, 1])

# adding slide with a few pandas dataframes with table titles
df2 = pd.DataFrame(np.array([[1, 2, 3], [4, 5, 6], [7, 8, 9]]), columns=['d', 'e', 'f'])
epp.addSlideWithTable(prs, 'Slide with a few pandas dataframes', [df1, df2])

# saving file
filename = "slide.pptx"
prs.save(filename)


```
### Installation
```
Copy file 'easy_python_pptx.py' to your work directory.
```


### List of functions:
```
addSlideTitle
   Adding new slide to presentation prs with title strTitle 

setSlideTitle
   Set style for slide's title with value text
   
addTextToSlide
   Adding text to slide

addListToSlide   
   Adding text or numeric list to slide
   
addTableToSlide
  Convert pandas dataframe to pptx table and insert it to slide with specific formats
  
addSlideWithList
  Adding slide with title and list of values
  
addSlideWithTable
  Adding slide with title and pandas dataframe (or dataframes)
```


