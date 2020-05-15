#Script that scrapes information about all of the available theses in fenix
#The information about these theses is then saved in a Theses.xlsx document in the working directory
#Very specific script
#Only works if fenix is in portuguese

from selenium import webdriver
import xlsxwriter

chromedriver = "/home/gnl/Documents/chromedriver"
driver = webdriver.Chrome(chromedriver)

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('Theses.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, 'ID')
worksheet.write(0, 1, 'Title')
worksheet.write(0, 2, 'Orienter')
worksheet.write(0, 3, 'Observations')
worksheet.write(0, 4, 'Requirements')
worksheet.write(0, 5, 'Goals')
worksheet.write(0, 6, 'Localization')
excel_line = 1;

driver.get("https://id.tecnico.ulisboa.pt/cas/login")

input("Enter the credentials, log in and after that press Enter to continue...")

students_tab_link = driver.find_element_by_link_text('Estudante')
students_tab_link.click()

dissertation_tab_link = driver.find_element_by_link_text('Candidatura a Dissertação')
dissertation_tab_link.click()

proposals_tab_link = driver.find_element_by_link_text('Propostas existentes')
proposals_tab_link.click()

# 5 pages to extract data from
for x in range(5):
	#all_results = driver.find_elements_by_class_name("sorting_1")
	all_results = driver.find_elements_by_tag_name("tr");
	# two first results are trash
	all_results.pop(0)
	all_results.pop(0)
	for line in all_results:
		cols = line.find_elements_by_tag_name("td");
		#details button
		description = line.find_element_by_css_selector("input[class='detailsButton btn btn-default']")

		worksheet.write(excel_line, 0, cols[0].text) #ID
		worksheet.write(excel_line, 1, cols[1].text) #Title
		worksheet.write(excel_line, 2, cols[2].text) #Orienter
		worksheet.write(excel_line, 3, description.get_attribute('data-observations'))
		worksheet.write(excel_line, 4, description.get_attribute('data-requirements'))
		worksheet.write(excel_line, 5, description.get_attribute('data-goals'))
		worksheet.write(excel_line, 6, description.get_attribute('data-localization'))
		excel_line = excel_line + 1
	if x!=4:
		driver.find_element_by_xpath('//a[@data-dt-idx="6"]').click() #click in the next page button
workbook.close()