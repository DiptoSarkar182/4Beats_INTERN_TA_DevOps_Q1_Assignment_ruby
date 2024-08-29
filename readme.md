# Web Scraping Script with Selenium and Excel

This project contains a Ruby script that reads keywords from an Excel file, performs Google searches for those keywords using Selenium WebDriver, and writes the results back to a new sheet in an updated Excel file.

## Prerequisites

Before you can run the script, you need to have the following installed:

1. [Ruby](https://www.ruby-lang.org/en/downloads/)
2. [Chrome Browser](https://www.google.com/chrome/)
3. [ChromeDriver](https://developer.chrome.com/docs/chromedriver/downloads)
4. Required Ruby gems:
    - `selenium-webdriver`
    - `roo`
    - `write_xlsx`
    - `fileutils`

## Installation

1. **Install Ruby:**
   Download and install Ruby from the official website: [Ruby Downloads](https://www.ruby-lang.org/en/downloads/).

2. **Install Chrome Browser:**
   Download and install the Chrome browser from the official website: [Google Chrome](https://www.google.com/chrome/).

3. **Install ChromeDriver:**
   Download the ChromeDriver that matches your Chrome browser version from the official website: [ChromeDriver Downloads](https://developer.chrome.com/docs/chromedriver/downloads).
    - Extract the downloaded file and place it in a directory included in your system's PATH.

4. **Install Required Gems:**
   Open a terminal or command prompt and run the following commands to install the required gems:
   ```sh
   gem install selenium-webdriver
   gem install roo
   gem install write_xlsx
   gem install fileutils

## Usage

1. Clone the repository: `git clone https://github.com/DiptoSarkar182/4Beats_INTERN_TA_DevOps_Q1_Assignment_ruby.git`
2. Navigate into the project directory: `cd 4Beats_INTERN_TA_DevOps_Q1_Assignment_ruby`
3. Run the Script: `ruby web_script.rb`
