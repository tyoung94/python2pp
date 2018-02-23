# python2pp

NOTE:
Clone this repo to your H drive and all paths will work

Whats been done:
 - Manipulate XML to put a series on the secondary y axis
 - Manipulate XML to convert a line series (recession usually) into an area chart
 
 What needs to be done:
 - Figure out a better way to store the variables (dictionaries?)
 - "IF" statements need to be implemented to check for existance of elements before creating them
 - Right now the python-pptx package can't write missing (NaN) values to PP.
    - potential workaround (convert all missing values to -9999 in python), then convert back to missingwith win32com
 
