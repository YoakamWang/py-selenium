import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
import pandas as pd


def initdriver():
    s = Service(r'C:\sources\edgedriver_win64\msedgedriver.exe')
    # driver = webdriver.Edge(executable_path=r'C:\sources\edgedriver_win64\msedgedriver.exe')
    # driver.get("https://www.baidu.com")
    wdriver = webdriver.Edge(service=s)
    wdriver.maximize_window()
    wdriver.get(
        "https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1")
    wdriver.get(
        "https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1")
    wwait = WebDriverWait(wdriver, 60)
    whome = wdriver.current_window_handle
    return wdriver, wwait, whome


# driver = initdriver()
# driver.maximize_window()
# driver.get(
#     "https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1")
# driver.get(
#     "https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1")
# wait = WebDriverWait(driver, 60)
# home = driver.current_window_handle


# driver.quit()

# items_identity = driver.find_elements(By.XPATH, "//*[@class='x-tree3-node-text']")
# insert exist frame //*[@id="ext-gen1"]/iframe[2]
# driver.find_element(By.XPATH, "//*[@class='x-grid3-scroller']//span[contains(text(),'" + 56 + "')]").click()


# buttons = driver.find_elements(By.XPATH, "//button[@class='x-btn-text ']")


# for insert new item
def newpage(driver, wait, home, parentnumber, cnumber, cname, ctype, csource, creversion, caccessory, csparepart,
            cgatheringpart,
            ccriticalcharac):
    driver.find_element(By.XPATH,
                        "//*[@class='x-grid3-scroller']//span[contains(text(),'" + parentnumber + "')]").click()
    driver.find_element(By.XPATH, "//button[contains(text(),'Insert New')]").click()

    wait.until(
        EC.number_of_windows_to_be(2)
    )
    toHandle = driver.window_handles  # CDwindow-BB52731DD9801849847E3F60C044733F
    # print(toHandle)
    for handle in toHandle:
        if handle == home:
            continue
        driver.switch_to.window(handle)
    driver.maximize_window()
    # print(driver.title)
    number = wait.until(EC.element_to_be_clickable((By.ID, 'number')))
    number.send_keys(cnumber)
    name = driver.find_element(By.XPATH,
                               "//input[@name='tcomp$attributesTable$OR:wt.part.WTPart:1152155555$___null!~objectHandle~partHandle~!_col_name___textbox']")
    name.send_keys(cname)
    assemblyMode = driver.find_element(By.ID, 'partType')
    Select(assemblyMode).select_by_value(ctype)
    source = driver.find_element(By.ID, "source")
    Select(source).select_by_value(csource)
    reverison = driver.find_element(By.XPATH, "//*[@onchange='PTC.revision.validateRevision(event);']")
    reverison.send_keys(creversion)
    # tracecode = driver.find_element(By.ID, "defaultTraceCode")
    # Select(tracecode).select_by_value("")
    # defaultunit = driver.find_element(By.ID, "defaultUnit")
    # Select(defaultunit).select_by_value("")
    # branditem = driver.find_element(By.ID, "BrndMrkItm")
    # Select(branditem).select_by_value("")
    accessory = driver.find_element(By.ID, "org.niladv.accessory")
    Select(accessory).select_by_value(caccessory)
    seperatepart = driver.find_element(By.ID, "org.niladv.SparePrt")
    Select(seperatepart).select_by_value(csparepart)

    gatheringparts = driver.find_elements(By.XPATH, "//*[@id='hidePartInStructure']//label")
    for part in gatheringparts:
        if part.text == cgatheringpart:
            part.click()
    material = driver.find_element(By.ID, "MATERIAL")
    material.location_once_scrolled_into_view
    criticalcharas = driver.find_elements(By.XPATH, "//*[@id='org.niladv.CriticalChar']//label")
    for chara in criticalcharas:
        if chara.text == ccriticalcharac:
            chara.click()
    finishButton = driver.find_element(By.ID, "ext-gen39")
    finishButton.click()
    wait.until(
        EC.number_of_windows_to_be(1)
    )
    driver.switch_to.window(home)


# for insert exist items
def existpage(driver, wait, parentnumbere, number):
    driver.find_element(By.XPATH,
                        "//*[@class='x-grid3-scroller']//span[contains(text(),'" + parentnumbere + "')]").click()
    driver.find_element(By.XPATH, "//button[contains(text(),'Insert Existing')]").click()
    existframe = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='ext-gen1']/iframe[2]")))
    # existframe = driver.find_element(By.XPATH, "//*[@id='ext-gen1']/iframe[2]")
    driver.switch_to.frame(existframe)
    numberex = driver.find_element(By.XPATH, "//label[contains(text(),'Number:')]/following-sibling::div//input")
    numberex.send_keys(number)
    searchbutton = driver.find_element(By.XPATH, "//button[contains(text(),'Search')]")
    searchbutton.click()
    okbutton = driver.find_element(By.XPATH, "//button[contains(text(),'OK')]")
    okbutton.click()
    wait.until(
        EC.visibility_of_element_located((By.ID, 'psbIFrame'))
    )
    # Switch to the frame //Switch to base frame  switchTo().defaultContent()
    # //*[@class="x-tree3-node-text"] to locate the items in the Identity
    driver.switch_to.frame("psbIFrame")


def getfile():
    path = sys.argv[1]
    return path


def main():
    mdriver, mwait, mhome = initdriver()
    filepath = getfile()
    with open(filepath, 'rb') as f:
        data = pd.read_excel(f, dtype=str)
        print(data.columns)
        print(data['Number'][0], data['Name'][0])

        Item = mwait.until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, data['Name'][0]))
        )
        Item.click()
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
        items_identity = mwait.until(
            EC.presence_of_all_elements_located((By.XPATH, "//*[@class='x-tree3-node-text']"))
        )

        for i in range(1, len(data['Level'])):
            # print(i, data['Level'][i])
            if int(data['Level'][i]) == int(data['Level'][i - 1]) + 1:
                # print(i, data['Level'][i])
                if data['Existing\n/New'][i] == 'Yes':
                    print('Yes')
                    existpage(mdriver, mwait, data['Number'][i - 1], data['Number'][i])
                    # run the insert exist page
                else:
                    print("No")
                    newpage(mdriver, mwait, mhome, data['Number'][i - 1], data['Number'][i], data['Name'][i],
                            data['Type'][i],
                            data['Source'][i],
                            data['Revision'][i], data['Accessory'][i], data['Spare\n Part'][i],
                            data['Gathering Part'][i],
                            data['Critical Characteristic'][i])
                    # run the insert new page i是当前行，i-1是父行
            elif int(data['Level'][i]) == int(data['Level'][i - 1]):
                # print(i, "dddd")
                for j in range(len(data['Level'][1:i]), -1, -1):
                    if int(data['Level'][j]) == int(data['Level'][i]) - 1:
                        # print(j, data['Number'][j])
                        if data['Existing\n/New'][i] == 'Yes':
                            print('Yes', i, j)
                            existpage(mdriver, mwait, data['Number'][j], data['Number'][i])
                            # run the insert exist page
                        else:
                            print("No", j)
                            # run the insert new page i是当前行，j是父行
                            newpage(mdriver, mwait, mhome, data['Number'][j], data['Number'][i], data['Name'][i],
                                    data['Type'][i],
                                    data['Source'][i],
                                    data['Revision'][i], data['Accessory'][i], data['Spare\n Part'][i],
                                    data['Gathering Part'][i],
                                    data['Critical Characteristic'][i])
                        break
            elif int(data['Level'][i]) < int(data['Level'][i - 1]):
                # print(i, "aaaa")
                for jj in range(len(data['Level'][1:i]), -1, -1):
                    if int(data['Level'][jj]) == int(data['Level'][i]) - 1:
                        # print(jj, data['Number'][jj])
                        if data['Existing\n/New'][i] == 'Yes':
                            print('Yes', i, jj)
                            # run the insert exist page
                            existpage(mdriver, mwait, data['Number'][jj], data['Number'][i])
                        else:
                            print("No", i, jj)
                            # run the insert new page, i是当前行，jj是父行
                            newpage(mdriver, mwait, mhome, data['Number'][jj], data['Number'][i], data['Name'][i],
                                    data['Type'][i],
                                    data['Source'][i],
                                    data['Revision'][i], data['Accessory'][i], data['Spare\n Part'][i],
                                    data['Gathering Part'][i],
                                    data['Critical Characteristic'][i])
                        break


if __name__ == "__main__":
    main()
"""Index(['Item', 'Number', 'Name', 'Qty', 'Unit', 'Level', 'Existing\n/New',
       'Revision', 'Critical Characteristic', 'Create as\n End Item', 'Type',
       'Assembly Mode', 'Source', 'Technical Part', 'Brand Marked Item',
       'Accessory', 'Spare\n Part', 'Gathering Part', 'Associate\n drawing',
       'Remark'],
      dtype='object')"""
# number=WebDriverWait(driver, 20).until(
#     EC.element_to_be_clickable((By.XPATH, "//*[@id='number']"))
# )
# https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1
# Part https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1
