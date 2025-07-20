# Scrapping-AFM-Results
A python code to scrap the students' results from the official website of E-Learning Unit that hosts the results.

**Author:** Shehab Habila – Biostatistician and R/Python Programmer   

---

## Overview

This Python script automates the process of retrieving student results (EGU, GIT, Skills, and Semester Total) - you can change it - from the official Alexandria Faculty of Medicine results page using Selenium and BeautifulSoup.

It reads a list of student IDs from an Excel file, queries the website for each ID, scrapes the grades, and appends the results into an output Excel file.

---

## Requirements

- Python 3.x
- Firefox browser
- [Geckodriver](https://github.com/mozilla/geckodriver/releases) (for Firefox control)
- Excel file with student IDs named `students_data.xlsx` and a column named `"ID"`

### Python Libraries

Install the required Python packages:

```bash
pip install selenium pandas bs4 os time
```

---

## Contact
- [LinkedIn](https://www.linkedin.com/in/shehab-habila/)
