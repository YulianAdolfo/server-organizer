**********************************************************************
**************************IMPORTANT*********************************

don´t delete this file, it´s very important
because the code takes it as reference to size the new excel file (rows and columns)

update the code with the new path:

fuction in code: func getColumnsDimension() 

dimension, _ := excelize.OpenFile("modelFormat/width-height.xlsx")
modelFormat is the folder