require 'selenium-webdriver'
require 'roo'
require 'write_xlsx'
require 'date'
require 'fileutils'

# Define the path to the updated Excel file
updated_excel_path = 'E:\TOP\assesment\4Beats_INTERN_TA_DevOps_Q1_Assignment_ruby\excel_file\updated_keywords.xlsx'

# Define the sheet name to read keywords from
input_sheet_name = 'sheet1'

# Open the Excel file for reading
excel_file = Roo::Spreadsheet.open('E:\TOP\assesment\4Beats_INTERN_TA_DevOps_Q1_Assignment_ruby\excel_file\keywords.xlsx')
sheet = excel_file.sheet(input_sheet_name)

# Get the current day and time to create a unique sheet name in the updated Excel file
today = Date.today.strftime("%A")
timestamp = Time.now.strftime("%H%M%S")
unique_sheet_name = "#{today}_#{timestamp}"

# Create a new workbook
new_workbook = WriteXLSX.new(updated_excel_path + '.tmp')

# Copy existing sheets to the new workbook if the file exists
if File.exist?(updated_excel_path)
  existing_workbook = Roo::Spreadsheet.open(updated_excel_path)
  existing_workbook.sheets.each do |sheet_name|
    existing_sheet = existing_workbook.sheet(sheet_name)
    new_sheet = new_workbook.add_worksheet(sheet_name)
    existing_sheet.each_with_index do |row, row_index|
      row.each_with_index do |cell, col_index|
        new_sheet.write(row_index, col_index, cell)
      end
    end
  end
end

# Add a new sheet with a unique name for today
worksheet = new_workbook.add_worksheet(unique_sheet_name)

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
  new_workbook.close

  # Replace the old workbook with the new one
  FileUtils.mv(updated_excel_path + '.tmp', updated_excel_path)

  # Close the browser
  driver.quit
end