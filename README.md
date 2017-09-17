<<<<<<< HEAD
### Description:
Most of the python MsOffice processing libraries, like openpyxk,python-docx,python-pptx,  are dealing with new office 2007 file formate (ie: xlsx,pptx,docx). This package can change old office 2003 to new,ie: doc2docx,xls2xlsx,ppt2pptx.
### PreCondition:
Pywin32 must be pre-installed and python3 is required.
### Tutorial:
**Usage is simple:**
## step1: tell converter where your data is: 
=======
## Description:
Most of the python MsOffice processing libraries, like openpyxk,python-docx,python-pptx,  are dealing with new office 2007 file formate (ie: xlsx,pptx,docx). This package can change old office 2003 to new,ie: doc2docx,xls2xlsx,ppt2pptx.
## PreCondition:
Pywin32 must be pre-installed and python3 is required.  
If you have difficulty on installing pywin32, go to [Christoph Gohlke](http://www.lfd.uci.edu/~gohlke/pythonlibs/) for wheel package,download whl file and pip install file name.
## Tutorial:
Usage is simple:
# step1: tell converter where your data is: 
>>>>>>> 348f977367797c8f6a2030ee83b7d3dee9b52195
`from changeOffice import Change`    
`c=Change("./data")`  
./data  is the root dir path you put your data in ,nested dirs works`
## step2: change formate and the api name is obvious:
`c.doc2docx()`   
`c.et2xls()# .et file must be converted to xls before  and then convert xls to xlsx`   
`c.xls2xlsx()`    
`c.ppt2pptx()`
<<<<<<< HEAD
## step3: to see the effect:
=======
# step3: to see the effect:
>>>>>>> 348f977367797c8f6a2030ee83b7d3dee9b52195
`print (c.get_allPath())`

