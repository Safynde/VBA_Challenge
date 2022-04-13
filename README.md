# VBA_Challenge
# Overview of Project: Explain the purpose of this analysis.
Steve’s parents are investing in green energy by buying DAQO New Energy Corp stocks. Steve wants to diversify the portfolio his parent’s portfolio by analyzing the stocks of 12 companies. The initial code works well for the analysis of the 12 stocks, but might not work efficiently if Steve decide to run thousands of stocks. Refactoring is a process of improving the code without creating a new functionality while leaving the program in working order. After building an initial VBA workbook to make the process easier, this challenge is about refactoring the initial VBA workbook into a clean code that is more efficient even with a larger size of data.	
# Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
In 2017, the average stock daily volume was 263,886,592 out of the 12 stocks with an average return of 67.3% only one stock had a negative return TERP (-7.2%).  Four company had a return above 100%, DQ (199%), SEDG (184.5%), ENPH (129.5%), FSLR (101.3%). In 2018, the average stock daily volume was 275,503,183 (An increase of 4.4%) with an average return of -8.5% (-75.5% decrease compared to 2017 only two stocks had a positive return RUN (84.4%) ENPH (81.9%) the positive return was less significant compared to 2017. The result between the original code and the refactored code stayed the same, therefore refactoring a code does not affect the results. Over the course of two years ENPH and RUN have a positive return.
![image](https://user-images.githubusercontent.com/101475984/163220555-1f30b383-cad8-4d3a-ba6f-ee81441be771.png)


## During the process of refactoring there are a couple of noticeable changes: 
### Creation of ticker 
The creation of a ticker index and creating three output arrays instead of initializing two variables for starting price and ending price. The ticker index is initialized to zero before looping start
![image](https://user-images.githubusercontent.com/101475984/163220382-0f4ebbc5-a460-4ca4-b3c3-d38f15e95c68.png)
![image](https://user-images.githubusercontent.com/101475984/163220442-f84db78c-e9a8-46ed-b8d9-cb63fbdd05ea.png)
### Loop
	A loop is created to initialize the ticker volume to zero instead of a loop through the ticker only. “i” is used in the refactored code instead of “j” from the initial code. In the “Green Stock” workbook, the loop is nested and only output the current data compare to the “VBA Challenge “workbook where only one loop moves the output to a separate loop, leading to a reduction in the number of steps in the refactored version of the code in “VBA Challenge “  workbook.
![image](https://user-images.githubusercontent.com/101475984/163220641-ff01cbfc-df7a-4e00-9fc7-805b347817a4.png)
![image](https://user-images.githubusercontent.com/101475984/163220672-c49ca5b8-d643-4bed-a0d2-7d52aa8ad1b1.png)

 ### Formating
In the Green Stock workbook, there is a button to format the data which creates an additional step compared to the “VBA Challenge “workbook.
 ![image](https://user-images.githubusercontent.com/101475984/163220718-091e9a9e-8bcf-4194-8328-8a2a379d10ca.png)
![image](https://user-images.githubusercontent.com/101475984/163220752-745bb813-ec9e-40fe-a4b6-dfb54c8a48fe.png)
### Runtime
	It takes longer to run the “Green Stock” workbook (36605.95 sec in 2018, 37483.23 sec in 2017) compared to the “VBA Challenges” workbook (0.204 sec in 2018, 0.188 sec for 2017). In the “Green Stock” workbook it takes longer to run 2017 stock compared to 2018. However, in the “VBA Challenges” workbook, it takes longer to run 2018 stocks compared to 2017.
![image](https://user-images.githubusercontent.com/101475984/163220832-e1093368-5e97-4b22-a7d9-84e4d1629936.png)
![image](https://user-images.githubusercontent.com/101475984/163220882-7c8ce082-10e0-4e19-a0e7-6ae09e45fe05.png)
![image](https://user-images.githubusercontent.com/101475984/163220918-277e4995-ca92-49b7-b635-eba931a26ca5.png)
![image](https://user-images.githubusercontent.com/101475984/163220948-39b4df37-b34d-4dbf-97ff-3e6ba2365581.png)
  
# The advantages and disadvantages of reactoring a code in general	
## What are the advantages or disadvantages of refactoring code?
The main benefits of refactoring a code are: 
Sustainability: since refactoring reduces the technical cost it makes the code easier to restore and recognize the code.
Helps finds bugs
Reduce the run time of a code 
Improves the design of the code which make it clean, easier to maintain and to understand	Minimizing or avoiding technical debt 
Keeps code up to date 
## The disadvantage of refactoring a code:
When the application is too big, the refactoring might not work as intended
When you do not have enough time and the delivery deadline is near it is risky to refactor the code
When the code is stable there is no need to refactor.
## <img width="383" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/101475984/163223687-77046624-685e-4ff6-9db5-f9ce45d62786.png">
<img width="386" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/101475984/163223721-23afbebc-7c4a-43e7-a5d2-6f25b5758f57.png">
How do these pros and cons apply to refactoring the original VBA script?
A loop is created to initialize the ticker volume to zero instead of a loop through the ticker only. “i” is used in the refactored code instead of “j” from the initial code. In the “Green Stock” workbook, the loop is nested and only output the current data compare to the “VBA Challenge” workbook where only the loop moves the output to a separate loop, leading to a reduction in the number of steps in the refactored version of the code in “VBA Challenge “workbook.
