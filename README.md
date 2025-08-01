# ModelMind: 
# AI-Powered Excel Data Financial Analyst
# Project Description

ModelMind is a web application built with Python and Flask that acts as an intelligent data analyst for your Excel files. 
Users can upload an Excel spreadsheet, ask a question in plain English, and receive a detailed answer generated by Google's Gemini AI or Chat-Gpt’s AI.
The application analyzes the uploaded data, and answers the user's question to provide insightful, context-aware responses.

#    Follow these instructions to set up and run the project on your local machine.
# Prerequisites:
Python 3.8 or newer <br/>

pip (Python package installer) <br/>

You will need to install several Python libraries using pip. The required packages are: <br/>
  * flask: The main web framework. <br/>
  * pandas: For reading and analyzing Excel data. <br/>
  * openpyxl: A dependency for pandas to handle .xlsx files. <br/>
  * google-generativeai: The official library for the Google Gemini API. <br/>
  * waitress: A production-quality server to run the app. <br/>
  * python-dotenv: To securely manage your API key. <br/>
  * werkzeug: A core dependency for Flask. <br/>
#   API Key Setup (.env)
The application requires a Google Gemini API key to function.
# Step 1: Get your API Key
  1. Go to the Google AI Studio.
  2. Sign in and click "Get API key" > "Create API key in new project".
  3. Copy the generated key.
# Step 2: Create the .env file
In the root of your project folder, create a new file named .env. <br/>
Open the .env file and add the following line, replacing your_api_key_here with the key you just copied: <br/>
GOOGLE_API_KEY='your_api_key_here' <br/>
#   Common Questions
# Personal Finance Management
* Analyze my spending for the last three months and categorize it. <br/>
* Based on my income and spending habits, what is a realistic monthly budget for me? <br/>
* Where are the biggest opportunities for me to save money each month? <br/>
* Track my subscriptions and identify any I might want to cancel. <br/>
* What is the most effective strategy for me to pay off my credit card debt? <br/>
# Investment Analysis and Strategy
* Analyze my current investment portfolio and identify areas of high risk. <br/>
* Based on my risk tolerance and financial goals, suggest a diversified portfolio of stocks and ETFs. <br/>
* What is the historical performance of my portfolio compared to the S&P 500? <br/>
* Provide a detailed financial analysis of [Company Name] stock, including key metrics and recent news. <br/>
* Compare the investment potential of gold versus real estate in the current economic climate. <br/>
# Market Trends and Economic Insights
* What are the current trends in the stock market? <br/>
* How is inflation expected to impact my investments and purchasing power? <br/>
* Provide an analysis of the current housing market in [Your City/State]. <br/>
* Explain the significance of the latest Federal Reserve interest rate decision. <br/>
* What does the current unemployment rate indicate about the health of the economy? <br/>
# Financial Education
* What is the difference between a stock and a bond? <br/>
* Explain how compound interest works with a clear example. <br/>
* What are the tax implications of short-term versus long-term capital gains? <br/>
* Define 'diversification' and explain why it is important for investing. <br/>

