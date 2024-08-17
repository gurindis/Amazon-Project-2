# Amazon-Project-2
**Summary:**  
bags.py (80+ lines of code) takes an excel file as an input, scrapes data from website and exports 3 different data points in excel format  

**Problem:**  
Amazon delivery station has many different indepedent delivery companies. Each delivery company has 30+ drivers working to deliver Amazon packages. Drivers load their vans with bags full of packages in the morning & they return those bags in the evening. Most drivers do not follow the SOP (Standard Operating Procedure) of returning their bags. SOP involves laying the empty bags individually on the carts (loose leaf), instead of folding 8-10 bags inside 1 bag  

**Solution:**  
This python script:  
1- reads an excel file, containing bag ids, as an input  
2- automates the process of going to Amazon website and searching for every bag id to  
3-get the name of drivers and the delivery company they work for  
4-BAG ID, DRIVER NAME, DELIVERY COMPANY NAME are exported to excel & the file is saved on the computer locally  

Excel file used as input contains bag ids that managers get by scanning the barcodes on every bag  

With over 200 bags returned improperly, it would take over 2 hours to pull this data manually every day  
But, this program runs in the background and only takes 1 minute to get it running  
