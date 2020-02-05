import keyboard  #for waiting for user input
import xlsxwriter
from time import sleep

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException

from secrets import username, password

driver = webdriver.Chrome()
driver.maximize_window()
####Login Process####
driver.get("https://angel.co/login")
email_in = driver.find_element_by_xpath("/html/body/div[1]/div[4]/div/div/div/div/div/div[1]/div[1]/form/input[4]")
email_in.send_keys(username)

pw_in = driver.find_element_by_xpath("/html/body/div[1]/div[4]/div/div/div/div/div/div[1]/div[1]/form/div[1]/input")
pw_in.send_keys(password)

login_btn = driver.find_element_by_xpath("/html/body/div[1]/div[4]/div/div/div/div/div/div[1]/div[1]/form/div[2]/input")
login_btn.click()
###Login Complete###

driver.get("https://angel.co/jobs")

### Waiting to set search conditions, press 'esc' to continue
keyboard.wait('esc')
############################################################

no_of_results = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[5]/div[2]/div/div/div[4]/h4")
print(no_of_results.text)
output = no_of_results.text

############################################################
no_of_jobs = int(output.strip(' results'))
################-------No_of_results-----------##########################
result = driver.find_elements_by_class_name("component_504ac")
###############################################
workbook = xlsxwriter.Workbook('jobs_description.xlsx')
worksheet = workbook.add_worksheet()
row = 1
col = 0

    ###############Inintialize Excel Sheet##########################
try:
    worksheet.write(0, 0, "company_name")
    worksheet.write(0, 1, "job_description")
    worksheet.write(0, 2, "compensation")
    worksheet.write(0, 3, "location")
    worksheet.write(0, 4, "job_type")
    worksheet.write(0, 5, "visa_sponsorship")
    worksheet.write(0, 6, "experience")
    worksheet.write(0, 7, "skills")
    #########################################
    limit = 5+no_of_jobs
    counter = 0
    for x in range(5,limit):
        counter+=1
        print(counter)
        front = "//*[@id='main']/div/div[5]/div[2]/div/div/div["
        end   = "]/div[2]/div/div[2]/button"
        path = front+str(x)+end

        ############################
        actions = ActionChains(driver)
        target = driver.find_element_by_xpath(path)
        actions.move_to_element(target)
        #############################
        test = driver.find_element_by_xpath(path)
        #sleep(10)
        try:
            test.click()
        except ElementClickInterceptedException:
                driver.execute_script("window.scrollTo(0,(%d*100);"%counter)
                try:
                    test.click()
                except ElementClickInterceptedException:
                    workbook.close()

        sleep(5)

        try:
            company_name = driver.find_element_by_css_selector(".\__halo_fontSizeMap_size--xl").text
        except NoSuchElementException:
            try:
                company_name = driver.find_element_by_xpath("/html/body/div[4]/div/div/div/div/div[1]/div[2]/h3").text
            except NoSuchElementException:
                company_name = driver.find_element_by_xpath("/html/body/div[3]/div/div/div/div/div[1]/div[2]/h3").text

        try:
            job_description = driver.find_element_by_css_selector(".\__halo_fontSizeMap_size--lg").text
        except NoSuchElementException:
            job_description = "None"
        try:
            compensation = driver.find_element_by_class_name("compensation_b6a2e").text
        except NoSuchElementException:
            compensation = "Missing"
        try:
            location = driver.find_element_by_class_name("location_a70ea").text
        except NoSuchElementException:
            location = "Not_Found"
        try:
            job_type = driver.find_element_by_css_selector(".characteristic_650ae:nth-child(2) > dd").text
        except NoSuchElementException:
            job_type = "Error"
        try:
            visa_sponsorship = driver.find_element_by_css_selector(".characteristic_650ae:nth-child(3) > dd").text
        except NoSuchElementException:
            visa_sponsorship = "Missing"
        try:
            experience = driver.find_element_by_xpath("//div[4]/dd").text
        except NoSuchElementException:
            experience = "None"
        try:
            skills = driver.find_element_by_css_selector(".characteristic_650ae:nth-child(5) > dd").text
        except NoSuchElementException:
            skills = "None"




    ####################################################################################################################

        try:
            cancel = driver.find_element_by_xpath("/html/body/div[4]/div/div/div/div/div[2]/div[6]/button[1]")
        except NoSuchElementException:
            try:
                cancel  =  driver.find_element_by_xpath("/html/body/div[10]/div/div/div/div/div[2]/div[6]/button[1]")
            except NoSuchElementException:
                try:
                    cancel  = driver.find_element_by_xpath("/html/body/div[6]/div/div/div/div/div[2]/div[6]/button[1]")
                except NoSuchElementException:
                    try:
                        cancel  =  driver.find_element_by_xpath("/html/body/div[3]/div/div/div/div/div[2]/div[6]/button[1]")
                    except NoSuchElementException:
                        cancel = driver.find_element_by_xpath("/html/body/div[7]/div/div/div/div/div[2]/div[6]/button[1]")
        cancel.click()
        ######################################################
        worksheet.write(row, col, company_name)
        worksheet.write(row, col+1, job_description)
        worksheet.write(row, col+2, compensation)
        worksheet.write(row, col+3, location)
        worksheet.write(row, col+4, job_type)
        worksheet.write(row, col+5, visa_sponsorship)
        worksheet.write(row, col+6, experience)
        worksheet.write(row, col+7, skills)
        row+=1

    workbook.close()
except:
    workbook.close()
