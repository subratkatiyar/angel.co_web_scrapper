import keyboard  #for waiting for user input
import xlsxwriter
from time import sleep

from selenium import webdriver
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

from secrets import username, password

options = webdriver.ChromeOptions()
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")
driver = webdriver.Chrome(options=options)

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

### Waiting to set search conditions, press 'esc' to continue and sleep for 5 seconds###
keyboard.wait('esc')
sleep(3)
############################################################
try:
    no_of_results = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[5]/div[2]/div/div/div[4]/h4")
except:
    sleep(3)
    try:
        no_of_results = driver.find_element_by_xpath("/html/body/div[1]/div/div/div[5]/div[2]/div/div/div[4]/h4")
    except NoSuchElementException:
        print("Try pressing esc after results load")


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
#opening_file/creating new file
file = open("job_requirements.txt", "a+", errors='ignore')
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
    worksheet.write(0, 8, "link")
    #worksheet.write(0, 9, "responsibilities")
    #########################################
    limit = 5+no_of_jobs
    counter = 0
    for x in range(5,limit):
        counter+=1
        print(counter)
        front = "//*[@id='main']/div/div[5]/div[2]/div/div/div["
        end   = "]/div[2]/div/div[2]/button"
        path = front+str(x)+end

        xpath_begin = "/html/body/div[1]/div/div/div[5]/div[2]/div/div/div["
        xpath_end = "]/div[2]/div/div[1]/a"
        xpath_url = xpath_begin+str(x)+xpath_end
        ############################
        actions = ActionChains(driver)
        target = driver.find_element_by_xpath(path)
        link = driver.find_element_by_xpath(xpath_url).get_attribute("href")

        actions.move_to_element(target)
        #############################
        apply = driver.find_element_by_xpath(path)
        try:
            apply.click()
        except NoSuchElementException:
            driver.execute_script("window.scrollTo(0,(%d*200));"%counter)
            try:
                apply.click()
            except NoSuchElementException:
                print("Uncaught exceptions")
                apply[counter-1].click()
        except ElementClickInterceptedException:
            driver.execute_script("window.scrollTo(0,(%d*250));"%counter)
            apply.click()

        sleep(2)

        try:
            company_name = driver.find_element_by_xpath("//*[@class='__halo_fontSizeMap_size--xl __halo_fontWeight_medium styles_component__1kg4S startupName_c5f67']").text
        except NoSuchElementException:
            sleep(10)
            try:
                company_name = driver.find_element_by_xpath("//*[@class='__halo_fontSizeMap_size--xl __halo_fontWeight_medium styles_component__1kg4S startupName_c5f67']").text
            except NoSuchElementException:
                print("Webpage is not loading")
                try:
                    sleep(10)
                    company_name = driver.find_element_by_xpath("//*[@class='__halo_fontSizeMap_size--xl __halo_fontWeight_medium styles_component__1kg4S startupName_c5f67']").text
                except NoSuchElementException:
                    print("Webpage is not loading")
                    try:
                        sleep(10)
                        company_name = driver.find_element_by_xpath("//*[@class='__halo_fontSizeMap_size--xl __halo_fontWeight_medium styles_component__1kg4S startupName_c5f67']").text
                    except NoSuchElementException:
                        print("Webpage is not loading")

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

        cancel = driver.find_element_by_xpath("//*[@class='styles_component__3A0_k styles_alternate__2u_Hm styles_regular__3b1-C component_21dbe']")

        try:
            cancel.click()
        except:
            driver.execute_script("return arguments[0].scrollIntoView();",cancel)
            try:
                cancel.click()
            except:
                raise KeyboardInterrupt

        #####################Get_Job_Information#################################
        main_window = driver.current_window_handle

        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[1])
        driver.get(link)
        try:
            responsibilities = driver.find_element_by_xpath("//*[@class='description_c90c4']").text
        except:
            sleep(5)
            try:
                responsibilities = driver.find_element_by_xpath("//*[@class='description_c90c4']").text
            except:
                sleep(10)
            try:
                responsibilities = driver.find_element_by_xpath("//*[@class='description_c90c4']").text
            except:
                sleep(5)




        driver.close()
        driver.switch_to.window(driver.window_handles[0])


        worksheet.write(row, col, company_name)
        worksheet.write(row, col+1, job_description)
        worksheet.write(row, col+2, compensation)
        worksheet.write(row, col+3, location)
        worksheet.write(row, col+4, job_type)
        worksheet.write(row, col+5, visa_sponsorship)
        worksheet.write(row, col+6, experience)
        worksheet.write(row, col+7, skills)
        worksheet.write(row, col+8, link)
        #worksheet.write(row, col+9, responsibilities)
        row+=1
        ###################Writing_the_description_to_file######################
        file.write("\n\n%d->%s\n%s\n%s"%(counter, company_name, job_description, responsibilities))



    file.close()
    workbook.close()

except KeyboardInterrupt:
    file.close()
    workbook.close()
except:
    file.close()
    workbook.close()
