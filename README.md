# Rutgers-JTS-HW2-VBA-Challenge
VBA Homework #2

The purpose of this exercise was to use the power of VBA to analyze hundred of thousands of lines in Stock Market data and return data charts with the following
  -Ticker Symbol 
  -Yearly Change 
  -Percentage Change 
  -Total Stock volume.
  
  The additional challenge was to sum the newly created chart even further by creating an additional chart listing the stock ticker symbol with the following values 
  -Greatest Percentage Increase 
  -Greatest Percentage Decrease 
  -Greatest Total Volume.
  
  The code begins by taking the total number of worksheets and creates a for k loop to complete the work on each individual sheet.
  The formatting of each sheet with headers, cell number formatting and column width are done.
  Next it finds the Lastline value to create an i loop that runs from the first line to the second to last line.
  The first ticker symbol, opening price & stock volume are saved.
  The i loop checks if the ticker symbol matches the ticker symbol the line below.  If it does then it adds the value to total volume and continues to the next i.
  If the ticker symbol does not match then it grabs the closing price, calculates the yearly change, percentage change & total stock volume.
  The stack value counter in increased and the new ticker symbol, new opeing price & new stock volume are set.
  This loop continues until the second to last line.  At that point the final closing price, total volume are added and yearly change, percentage change & total stock volume are set.
  The next i loop goes through the new chart and pulls the greater % increase, greater % decrease & total volume by comparing each value and when a higher or lower value is found the new value is updated until all the data has been compared.
  The last step is to create "conditional formatting" in the Yearly changed column by using a For each loop.  If the value is less that 0, color the cell interior red.  If the value is greater that 0, color the cell interior green.
  The code ends by going to the next k which is the next sheet.
  
  
  
