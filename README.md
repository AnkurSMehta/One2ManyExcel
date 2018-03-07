# One2ManyExcel
R / Python code to easily generate multiple excel files from a single file
Very often, business users store structured data in excel files and then need to filter portions of the data and send it out to multiple people.
As an example, data on potential customers in a single excel file may need to be split into multiple data sheets or files based on certain criteria, e.g. geography; each individual data file pertaining to a single geographical unit may need to be sent to a salesperson or customer service administration representative for that geography
This R code reads in excel or csv files, filters it using criteria in the data and builds out multiple excel files based on the criteria
Dependencies - dplyr, xlsx packages
