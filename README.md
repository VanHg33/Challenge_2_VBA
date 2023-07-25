# VBA-Challenge
Module 2 Challenge

# **Background**

  You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyse    generated stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the bonus       Challenge tasks.

# **Instructions**

  Create a script that loops through all the stocks for one year and outputs the following information:
  - The ticker symbol
  - Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
  - The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
  - The total stock volume of the stock.
  - NOTE: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

**Bonus**

  - Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total           volume".  
  - Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

**Other Considerations**

  - Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your       code should run on this file in under 3 to 5 minutes.
  - Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with      the click of a button.
  - Some assignments, like this one, contain a bonus. It is possible to achieve proficiency for this assignment without completing the        bonus, but the bonus is an opportunity to further develop your skills and receive extra points for doing so.

Steps by steps to read through the code file:
1. Loop through all the worksheets using For each 
2. Create all the column headers that the Challenge requires
3. For big dataset, define Lastrow 
4. Create all the variables that use in this code
5. Creat For loop to loop through each row of the worksheet, and make the calculations as required. 
6. Using If inside For loop to let the code loop through all the values, to find the createst increase, greatest decrease and creates volumn.
7. Apply conditional formatting for both yearly change column and percent change column as required. THe code are referenced from website VBA Conditional Formatting (Sethi, T, 29th June 2023, VBA Conditional Fomatting in Excel, https://www.wallstreetmojo.com/vba-conditional-formatting/)
8. Apply formatting of percentage for needed values. 
