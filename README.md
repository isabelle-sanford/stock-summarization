# stock-summarization
VBA code for summarizing a sheet of stocks, & accompanying example pics

This code is meant to run on an Excel workbook containing one or more sheets of stock listings. These listings should be in the format of the first seven columns as follows: 
A: Ticker
B: Date [UNUSED]
C: Open
D: High [UNUSED]
E: Low [UNUSED]
F: Close
G: Volume

It will output a summary table to the right of these columns consisting of the ticker, the change from first open to last close of a particular stock (reading from the first and last row listing that particular ticker; stocks must be in order by date for this to be effective), the percentage change (colored green or red depending if it's positive or negative), and the total volume over the listed dates. 
There is also a summary of that table even further to the right, which gives values for the ticker with: the highest percentage change, the lowest percentage change, and the highest volume.

Screenshots are also included to provide a visual example. 

This code was made as part of the VBA unit at USC Viterbi Boot Camp. 
