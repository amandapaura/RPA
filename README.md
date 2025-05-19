# RPA

🔗 Power BI Report Downloader

This project automates the process of extracting information from Power BI workspaces and downloading reports using the Power BI API and Selenium.


📁 Project Structure

api_powerbi.py – Class with methods to connect to the Power BI API (authentication).

pbi_workspaces_reports.py – Connects to the API and exports a JSON with workspace and report details (including URLs).

pbi_web_scraping.py – Uses Selenium to open report URLs and download the PBIX files via web scraping.

main.py - execute json collection and scraper


🛠️ Features

Connects to the Power BI REST API using msal.

Exports metadata (workspace and report names, IDs, and URLs) to a JSON file.

Automates the PBIX download process via browser interaction.


✅ Requirements

Python 3.8+

MSAL

Requests

Selenium

ChromeDriver

Keyring 


⚠️ Make sure you have access permissions to the Power BI service and reports.
