# project2
## "Sport Data Analysis"
## Project Overview
This project was completed as part of a real-world Excel case study for XYZ Co Pvt Ltd — a company that organizes international sports events. The task involved cleaning, standardizing, analyzing, and reporting on a sportsmen membership dataset covering 50 athletes from 11 countries across multiple sports disciplines.
The project was completed in 3 stages: Data Cleaning, Data Analysis, and Report Generation.
## Tools & Technologies Used
- Microsoft Excel
  - Index and Match
  - IF, COUNTIF, CONCATENATE Functions
  - Text Functions (UPPER, LOWER, LEFT, RIGHT)
  - Pivot Tables
  - Data Formatting and Conditional Formatting
  - Slicers and Filters
  - Custom Number Formatting
## 📂 Dataset Details
- Total Records:50 Sportsmen
- Countries Covered: 11 (USA, UK, Australia, Germany, France, Argentina, Spain, Netherlands, Austria, Brazil, Sweden)
- Languages: English, German, French, Spanish, Portuguese, Dutch, Swedish
- Sports Types:Indoor and Outdoor
- Key Columns:Member ID, Full Name, Birthdate, Gender, Country, Language, Email, Weight, Sport, Salary
## 🔍 Project Stages
### ✅ Stage 1 — Data Cleaning & Standardization
**Task 1: Standardizing the Dataset**
- Full Name in prescribed format: `PREFIX FIRSTNAME LASTNAME` (All Uppercase)
- Country Name using VLOOKUP from Location sheet
- Language spoken by each sportsman using Location sheet
- Email Addresses based on language:
  - English speakers → `lastname.firstname@xyz.org`
  - All others → `lastname.firstname@xyz.com`
- Populated Sport Location (Indoor/Outdoor) using SPORT sheet
Task 2: Data Formatting
- Formatted Member ID as always 3 digits (001, 002 ... 050)
- Formatted Birthdate as `dd mmm' yyyy` (Example: 09 May' 1986)
- Added units to Weight column (Example: 80 kg)
- Formatted Salary in thousands with conditional decimal places:
  - Below ₹1,00,000 → 2 decimal places (Example: 87.67 k)
  - Above ₹1,00,000 → 1 decimal place (Example: 123.2 k)
### ✅ Stage 2 — Data Analysis
Task 1: Pivot Table Summary
- Created Pivot Table showing Count of Athletes by Country and Gender
- Rows: Country Names
- Columns: Gender (Male / Female)
- Grand Totals removed for clean view
Task 2: Excel Functions Summary
- Built cross table using Excel functions (COUNTIFS)
- Extracted distinct Gender values using Remove Duplicates + Transpose
- Extracted distinct Country Names
- Calculated count of athletes per Country and Gender combination
### ✅ Stage 3 — Report Generation
- Created Pivot Table Report in tabular layout showing:
  - Member ID, Full Name, Email, Gender
  - Year of Birth, Country Name, Language, Sport
- Added Slicer to filter report by Sport Location (Indoor / Outdoor)
- Removed expand/collapse buttons for clean presentation
- Removed Grand Totals
## 💡 Key Insights Found
- 🇫🇷 France had the highest number of athletes — 9 sportsmen (6 Male, 3 Female)
- 🇦🇺 Australia had the most female athletes — 6 Female sportsmen
- 🇧🇷 Brazil had only Male athletes in this dataset
- 🇪🇸 Spain had only Female athletes in this dataset
- Most popular sports included Cycling Road, Beach Volleyball, Triathlon and Athletics
- English was the most spoken language among athletes
- Sports were evenly split between Indoor and Outdoor venues
- Salary formatting revealed wide range — from **10.2k to 117.4k
## 📸 Dashboard Screenshots
<img width="320" height="266" alt="image" src="https://github.com/user-attachments/assets/b31b5fea-f9f2-43bd-9c0f-897aaf9848a4" />
## 📁 Files in This Repository
- `Excel_case_study_1.xlsx` — Complete Excel file with all 3 stages
- `README.md` — Project documentation
- `dashboard_screenshot.png` — Screenshot of final report
## 🚀 How to View This Project
1. Download the `.xlsx` file
2. Open in Microsoft Excel
3. Navigate through sheets:
   - `SPORTSMEN` — Cleaned and formatted data
   - `ANALYSIS` — Pivot table and summary analysis
   - `REPORT` — Final pivot report with slicer
4. Use the Sport Location slicer to filter between Indoor and Outdoor sports
## 💼 Skills Demonstrated
- Data Cleaning and Standardization
- Advanced Excel Formulas (Index & match, IF, COUNTIFS, TEXT)
- Custom Number and Date Formatting
- Pivot Table Creation and Formatting
- Slicer and Filter Implementation
- Cross Table Analysis using Excel Functions
- Data Reporting and Presentation
## Contact
- 🔗 LinkedIn: https://www.linkedin.com/in/manish-k-jha-80795b284/
- 📧 Email: manish05jha04@gmail.com
