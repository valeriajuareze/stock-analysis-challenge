# Stock Analysis Challenge
## Overview of Project

Steve´s parents are passionate in green energy and they are planning to invest in just one green company, but Steve does not agree at all. So, he wants me to help him to analize the databases of all green stocks energy to find the best companies to invest in. Therefore, the purpose of this project is to analyze all the stocks with VBA, to reduce the chances of accidents and investment errors.

## Results 
### Refactoring Code
First, to create my code, what I did was create the ticker Index with a variable 0, this, for everytime it changes of ticker start again in 0. Then, I created the three output arrays, this was important because if I did them as variables, then my code will not be runned, so I putted them: (12) to make them arrays instead of variables.

![image](https://user-images.githubusercontent.com/108365182/179141745-eb13aa5f-a10a-4550-9484-47b5b6007a4e.png)

Next, I created a loop to initialize the tickerVolume in 0, and I ended it because I wanted that everytime it will changes of ticker, starts in 0. For this step, I used the function For iteretator to change the value over the course. Also, here I started another For loop to start finding new values.

![image](https://user-images.githubusercontent.com/108365182/179142626-b12e188c-3fa0-461c-ab7b-74d43d6f1f2d.png)

The main purpose of creating a loop throughout all the rows in the datasheets was to repeat it until it finds the volume, the starting and ending prices of each ticker. Here, I used the conditionals and also the boolean operators which were quite useful to find the first and last cell of each ticker. 

![image](https://user-images.githubusercontent.com/108365182/179143795-c22660ea-96be-497a-8ccf-9b19357ca75f.png)

Finally, I activated the All Stocks Analysis sheet to report all the findings. I putted again the iterator with the main purpose to report all of the tickers in the chart. 

![image](https://user-images.githubusercontent.com/108365182/179144594-1b781e46-77f0-4aaf-8cce-6b7b71badb2e.png)

### Results of Refactoring Code from 2018

In the original code from 2018 we had an execution time of:

![image](https://user-images.githubusercontent.com/108365182/179145706-8d45b36c-9ab5-406f-8161-2fd78a6f540a.png)

With The following results :

![image](https://user-images.githubusercontent.com/108365182/179148821-8bf13bcf-3184-404d-b844-19fe2b2ed1b4.png)

On the other hand in the refactoring code form 2018, we had an execution time of:

![image](https://user-images.githubusercontent.com/108365182/179148616-018a4987-353a-450c-b75e-5ea307715318.png)

With the same results:

![image](https://user-images.githubusercontent.com/108365182/179146085-7b859a57-71e6-4b32-8c7e-618ff603b123.png)

### Results of Refactoring Code from 2017 

In the original code from 2017 we had an execution time of:

![image](https://user-images.githubusercontent.com/108365182/179146347-5853fddd-7f88-4f25-86bf-c13aaedc1ea2.png)

With the following results: 

![image](https://user-images.githubusercontent.com/108365182/179148314-dbb5a68d-7a81-4dc3-a352-c3484eb5e826.png)

On the other hand in the refactoring code form 2017, we had an execution time of:

![image](https://user-images.githubusercontent.com/108365182/179146468-605b0d70-3f04-467c-a9c8-cb65c7b3c582.png)

With the same results:

![image](https://user-images.githubusercontent.com/108365182/179146565-9107a019-2e7f-4ed6-a0ad-edd10341f307.png)

### Recommendation 

So, in base of the previous results, we can deduce that it will be a bad choice for Steve´s parents to invest in DAQO company, instead, I recommend them to invest in two companies, which are ENPH and RUN. We can say that in 2017 was a good year to invest in DAQO company, but according to 2018, we can say that it was not a good year for the company, so we recommend them to invest in other companies.

## Summary

### What are the advantages or disadvantages of refactoring code?

One of most important advantages of refactoring code is to make it more efficient and faster. also, to do it more readable to next users with the main purpose to evolve it in the future. Otherwise, the disadvantages exist, one of them is that we could create errors into our code instead of make it easier. 

### How do these pros and cons apply to refactoring the original VBA script?

During the refactoring of the original VBA script, I had a lot of errors declarating in different way the arrays and outputing the results, which made the process more difficult to me and made me erase a lot of my scripts. But finally, I reduced more than a half of the execution time, making them more efficient and faster that the original one. Sincerely, I think that my disadvantage is that I did not make the code more readable than the original one, because I reduced a few steps, but at least I think it is comprehensible for the readers. 

