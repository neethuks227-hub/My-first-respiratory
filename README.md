# My-first-repositary
Assignment 1 Data exploration
Given  dataset contains information about various products, including:Product ID,Product Name ,Brand name,Price ,Quantity & Categorey .Here we have to do the Excel Functions 
**1) Sum, Count, Average:**		
Total Price of All Products: =SUM(D2:D22)
Total Count of Products: =COUNT(D2:D22)
Average Price of Products: =AVERAGE(D2:D22)
**2) Min and Max:	**
Minimum Price**: =MIN(D2:D22)
Maximum Price**: =MAX(D2:D22)
	**3) IF Function:**											
		• Using an IF function, create a new column named Price Range to categorize products with a price greater than or equal to $500 as 'High Price' and others as 'Standard Price'.	
    =IF(D2>=500, "High price", "standard price")
 **   4) SUMIF and COUNTIF:		**		
	• Calculate the total price for products in the 'Electronics' category using the SUMIF function.	
            =SUMIF(F2:F22, "Electronics", D2:D22)
	• Determine the count of products with a price less than $100 using the COUNTIF function.					
            =COUNTIF(D2:D22, "<100")   
    **5) Text Formatting - LEFT, RIGHT, MID:	**								
	• Create a new column named Day with the first 2 characters of each 'Product ID' using the LEFT function.		
                =LEFT(A2, 2)
	• Create a new column named Country Code by extracting the last 2 characters from the 'Product ID' column using the RIGHT function.	
                =RIGHT(A2, 2)
	• Create a new column named Month by extracting 4th to 6th characters from the 'Product ID' column using the MID function.	
                =MID(A2, 4, 3)
									
									
