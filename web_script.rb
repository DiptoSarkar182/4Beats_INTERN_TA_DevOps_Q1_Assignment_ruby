require 'selenium-webdriver'
require 'roo'
require 'write_xlsx'
require 'date'
require 'fileutils'

# Define the path to the updated Excel file
updated_excel_path = 'E:/TOP/assesment/devops_intern_ruby/excel_file/updated_keywords.xlsx'

# Define the sheet name to read keywords from
input_sheet_name = 'sheet1'

# Open the Excel file for reading
excel_file = Roo::Spreadsheet.open('E:/TOP/assesment/devops_intern_ruby/excel_file/keywords.xlsx')
sheet = excel_file.sheet(input_sheet_name)

# Get the current day to create a new sheet in the updated Excel file
today = Date.today.strftime("%A")

# Check if the updated Excel file already exists
if File.exist?(updated_excel_path)
  # Open the existing file and add a new sheet for today
  workbook = WriteXLSX.new(updated_excel_path)
  worksheet = workbook.add_worksheet(today)
else
  # Create a new Excel file for writing
  workbook = WriteXLSX.new(updated_excel_path)
  worksheet = workbook.add_worksheet(today)
end

# Read the header row to determine column indices
header = sheet.row(1)
keyword_index = header.index('Keywords')
longest_option_index = header.index('Longest Option')
shortest_option_index = header.index('Shortest Option')

# Copy the header row to the new sheet
header.each_with_index do |header_cell, index|
  worksheet.write(0, index, header_cell)
end

# Initialize Selenium WebDriver
driver = Selenium::WebDriver.for :chrome

begin
  # Iterate through each row in the sheet
  sheet.each_row_streaming(offset: 1).with_index(1) do |row, row_index|
    keyword = row[keyword_index]&.cell_value&.strip
    next if keyword.nil? || keyword.empty?

    begin
      driver.navigate.to "https://www.google.com/search?q=#{keyword}"

      # Wait for the search results to load
      wait = Selenium::WebDriver::Wait.new(timeout: 20) # Increased timeout to 20 seconds
      wait.until { driver.find_element(css: 'h3') }

      # Extract search results
      results = driver.find_elements(css: 'h3').map(&:text).reject(&:empty?)

      # Filter out non-English results (optional)
      results = results.select { |result| result.ascii_only? }

      # Debugging: Print the results to ensure they are being captured
      puts "Keyword: #{keyword}"
      puts "Results: #{results.inspect}"

      # Find longest and shortest options
      longest_option = results.max_by(&:length)
      shortest_option = results.min_by(&:length)

      # Debugging: Print the longest and shortest options
      puts "Longest Option: #{longest_option}"
      puts "Shortest Option: #{shortest_option}"

      # Write the original keyword and the results back to the new Excel file
      worksheet.write(row_index, keyword_index, keyword)
      worksheet.write(row_index, longest_option_index, longest_option)
      worksheet.write(row_index, shortest_option_index, shortest_option)
    rescue Selenium::WebDriver::Error::TimeoutError
      puts "Timeout while searching for keyword: #{keyword}"
      worksheet.write(row_index, keyword_index, keyword)
      worksheet.write(row_index, longest_option_index, "Timeout")
      worksheet.write(row_index, shortest_option_index, "Timeout")
    rescue StandardError => e
      puts "An error occurred: #{e.message}"
      worksheet.write(row_index, keyword_index, keyword)
      worksheet.write(row_index, longest_option_index, "Error")
      worksheet.write(row_index, shortest_option_index, "Error")
    end
  end
ensure
  # Close the workbook
  workbook.close

  # Close the browser
  driver.quit
end