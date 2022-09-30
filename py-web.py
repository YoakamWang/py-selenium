from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
# edriver = "C:\sources\edgedriver_win32"
import time
from selenium.webdriver.support.wait import WebDriverWait

s = Service(r'C:\sources\edgedriver_win64\msedgedriver.exe')
# driver = webdriver.Edge(executable_path=r'C:\sources\edgedriver_win64\msedgedriver.exe')
# driver.get("https://www.baidu.com")
driver = webdriver.Edge(service=s)
driver.maximize_window()
driver.get(
    "https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1")
driver.get(
    "https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1")
# driver.quit()
wait = WebDriverWait(driver, 40)
Item = wait(
    EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, 'SCRUB SYSTEM-48 IN DISK'))
)
Item.click()
structure = wait.until(
    EC.element_to_be_clickable((By.ID, 'infoPageinfoPanelID__infoPage_myTab_psb_productStructureGWT'))
)
structure.click()
# wait the frame available to switch
wait.until(
    EC.visibility_of_element_located((By.ID, 'psbIFrame'))
)
# Switch to the frame //Switch to base frame  switchTo().defaultContent()
# //*[@class="x-tree3-node-text"] to locate the items in the Identity
driver.switch_to.frame("psbIFrame")
time.sleep(10)  # wait for all items load complete currently, will find a better way latter.
items_identity = wait.until(
    EC.presence_of_all_elements_located((By.XPATH, "//*[@class='x-tree3-node-text']"))
)
# items_identity = driver.find_elements(By.XPATH, "//*[@class='x-tree3-node-text']")
# insert exist frame //*[@id="ext-gen1"]/iframe[2]
home = driver.current_window_handle
for it in items_identity:
    if "56414E19" in it.text:
        it.click()
        buttons = driver.find_elements(By.XPATH, "//button[@class='x-btn-text ']")
        for button in buttons:
            if button.text == "Insert New":
                button.click()

wait.until(
    EC.number_of_windows_to_be(2)
)
toHandle = driver.window_handles  # CDwindow-BB52731DD9801849847E3F60C044733F
print(toHandle)
for handle in toHandle:
    if handle == home:
        continue
    driver.switch_to.window(handle)
driver.maximize_window()
print(driver.title)
# number=WebDriverWait(driver, 20).until(
#     EC.element_to_be_clickable((By.XPATH, "//*[@id='number']"))
# )
# https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1
# Part https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1
