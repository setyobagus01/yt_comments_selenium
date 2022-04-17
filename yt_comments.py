from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import xlsxwriter

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# put your youtube url here
driver.get("https://www.youtube.com/watch?v=2qUCyW7ewPs")

actions = ActionChains(driver)

# excel filename
workbook = xlsxwriter.Workbook("yt_comments.xlsx")

worksheet = workbook.add_worksheet()
cell_format = workbook.add_format()
cell_format.set_bold()
# header 
worksheet.write(0, 0, "comments", cell_format)


# set how long it takes to scroll
for i in range(500):
    actions.scroll(0, 0, 0, 100) # scroll to y-direction by 100 for each iteration
    actions.perform()

comments = driver.find_elements(By.XPATH, "//yt-formatted-string[@id='content-text' and @slot='content' and @class='style-scope ytd-comment-renderer']")
for index, comment in enumerate(comments):
    worksheet.write(index + 1, 0, comment.text)

workbook.close()