🔍 Project Overview

This project demonstrates how to build an offline, ChatGPT-like analytical system using only Excel formulas.
It allows users to ask natural language business questions (e.g., “top 5 average budgets for projects last month”) and instantly generates accurate financial reports — without using cloud-based AI tools.
The solution is inspired by Chandoo’s approach to leveraging advanced Excel features for secure, reliable, and verifiable data analysis.

🚨 Problem with Using AI Tools in Finance

Many organizations hesitate to use tools like ChatGPT or Copilot for financial analysis due to:
Data Privacy Risks – Sensitive finance data is exposed to external cloud platforms
High Cost – AI subscriptions and usage credits increase operational costs
Accuracy Issues – Financial data has no scope for ambiguity, and AI models may hallucinate or misinterpret numbers

💡 Solution: Excel GPT (Offline & Verifiable)

This project solves the above challenges by creating a logic-driven analytics engine inside Excel.

Key Capabilities

Accepts plain English prompts directly in Excel
Converts user intent into structured queries using formulas
Generates accurate, auditable, and repeatable results
Keeps all data 100% local and secure

🛠️ How the System Works

The system processes 13,000+ rows of financial data (Actuals, Budgets, Forecasts) using the following workflow:
Prompt Tokenization
Identifies keywords (sum, average, project, business unit, YTD, etc.)
Uses predefined synonym lists for measures and dimensions
Keyword Matching (XLOOKUP)
Matches user input with known business terms
Determines what calculation and grouping is required
Numeric Extraction (REGEXTRACT)
Extracts values like “5” from prompts such as “top 5 projects”
Time-Based Filtering
Dynamically applies filters like This Month, Last Month, or YTD
Data Aggregation (GROUPBY)
Summarizes and sorts data
Applies ranking logic for top/bottom results

🧰 Tools & Excel Features Used

Advanced Excel Functions
XLOOKUP, LET, GROUPBY, SWITCH, REGEXTRACT
Excel Labs Add-in
Advanced Formula Editor for managing complex formulas
Logic-Based Design
Fully deterministic and auditable calculations

📈 Example Queries Supported

Total actuals by expense type this month
Top 5 average budgets for projects last month
Max program actuals YTD
Each result can be manually verified using Excel’s status bar, ensuring 100% accuracy.

🔄 Alternative Approach: Pivot Tables

For users who prefer a simpler solution, the project also highlights Pivot Tables as a built-in alternative:
Drag-and-drop interface
Fully offline and private
No subscriptions required
Ideal for fast exploratory analysis

🎯 Business & Data Analytics Value

Improves decision-making speed
Enhances stakeholder trust in numbers
Eliminates AI hallucination risks
Reduces manual reporting effort
Ensures data privacy and auditability

📌 Use Cases

Financial Reporting
Budget vs Actual Analysis
Management Dashboards
Secure Enterprise Reporting
Offline Analytics Environments

👤 Author

Chakali Sasi Kumar
Aspiring Data Analyst | Business Analyst
Skills: Excel, SQL, Python, Power BI, Data Analysis

📌 If you found this project useful, feel free to star ⭐ the repository!
