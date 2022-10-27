import time

from selenium.common import TimeoutException
from selenium.webdriver.support.wait import WebDriverWait

import py_web_sele as se
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC

mdriver, mwait, mhome = se.initdriver()
item = mwait.until(
    EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'SCRUB SYSTEM-48 IN DISK')))
item.click()
structure = mwait.until(
    EC.element_to_be_clickable((By.ID, 'infoPageinfoPanelID__infoPage_myTab_psb_productStructureGWT'))
)
structure.click()
# wait the frame available to switch
mwait.until(
    EC.visibility_of_element_located((By.ID, 'psbIFrame'))
)
# Switch to the frame //Switch to base frame  switchTo().defaultContent()
# //*[@class="x-tree3-node-text"] to locate the items in the Identity
mdriver.switch_to.frame("psbIFrame")
time.sleep(10)  # wait for all items load complete currently, will find a better way latter.
mwait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//*[@class='x-tree3-node-text']"))
)
se.click_parent(mdriver, parentnumber='56414E19')
try:
    WebDriverWait(mdriver,4).until(EC.element_attribute_to_include((By.XPATH, "//button[contains(text(),'Insert New')]"),'disabled'))
except TimeoutException:
    print("ke yi dian ji")
else:
    print('dian ji bu liao')
# tests=mwait.until(EC.visibility_of_all_elements_located((By.XPATH,"//*[@class='cat-fit-match']")))
# if len(tests)>0:
# insert_new_button = mwait.until(
#     EC.visibility_of_element_located((By.XPATH, "//button[contains(text(),'Insert Existing')]")))
# new_button_state = insert_new_button.get_attribute('aria-disabled')
# insert_new_button.click()
# print(new_button_state)
