import sys
from selenium import webdriver
from selenium.common import NoSuchWindowException, StaleElementReferenceException, ElementClickInterceptedException, \
    TimeoutException, NoSuchElementException, NoAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.select import Select
import pandas as pd
from selenium.webdriver import Keys


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


def click_parent(driver, parentnumber):
    # all_links = wait.until(EC.visibility_of_all_elements_located((By.CLASS_NAME, 'x-tree3-node-text')))
    # print(len(all_links))
    # for link in all_links:
    #     text = link.find_element(By.XPATH, "./span").text
    #     if parentnumber in text:
    #         # driver.execute_script("arguments[0].click();", link) 有bug，之点击第一个
    #         try:
    #             link.click()
    #         except StaleElementReferenceException:
    #             wait.until(EC.element_to_be_clickable((By.XPATH,
    #                                                    "//*[@class='x-grid3-scroller']//span[contains(text(),'" + parentnumber + "')]")))
    #             driver.find_element(By.XPATH,
    #                                 "//*[@class='x-grid3-scroller']//span[contains(text(),'" + parentnumber + "')]").click()
    search_parent = driver.find_element(By.XPATH, "//*[@class=' x-form-field-wrap  x-component x-border']/input")
    search_parent.clear()
    time.sleep(0.5)
    search_parent.send_keys(parentnumber)
    time.sleep(0.5)
    search_parent.send_keys(Keys.ENTER)


# for insert new item
def newpage(driver, wait, home, parentnumber, cnumber, cname, ctype, csource, creversion, caccessory,
            csparepart,
            cgatheringpart,
            ccriticalcharac, cqty, ctechpart):
    try:
        click_parent(driver, parentnumber)
    except NoSuchElementException:
        js = 'document.getElementsByClassName("x-grid3-scroller")[0].scrollTop=300'
        driver.execute_script(js)
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, "//*[@class='x-grid3-scroller']//span[contains(text(),'" + parentnumber + "')]")))
    try:
        WebDriverWait(driver, 4).until(
            EC.element_attribute_to_include((By.XPATH, "//button[contains(text(),'Insert New')]"), 'disabled'))
    except TimeoutException:
        try:
            insert_new_button = wait.until(
                EC.visibility_of_element_located((By.XPATH, "//button[contains(text(),'Insert New')]")))
            insert_new_button.click()
        except ElementClickInterceptedException:
            insert_new_button = wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Insert New')]")))
            driver.execute_script("arguments[0].click();", insert_new_button)
        wait.until(
            EC.number_of_windows_to_be(2)
        )
        tohandle = driver.window_handles  # CDwindow-BB52731DD9801849847E3F60C044733F
        # print(toHandle)
        for handle in tohandle:
            if handle == home:
                continue
            driver.switch_to.window(handle)
        driver.maximize_window()
        number = wait.until(EC.element_to_be_clickable((By.ID, 'number')))
        number.send_keys(cnumber)
        name = driver.find_element(By.XPATH, "//td[@attrid='name']/input[@type='text']")
        name.send_keys(cname)
        assemblymode = driver.find_element(By.ID, 'partType')
        Select(assemblymode).select_by_value(ctype)
        source = driver.find_element(By.ID, "source")
        Select(source).select_by_value(csource)
        reverison = driver.find_element(By.XPATH, "//*[@onchange='PTC.revision.validateRevision(event);']")
        reverison.clear()
        try:
            wait.until(EC.alert_is_present())
            driver.switch_to.alert.accept()
        finally:
            if pd.isna(creversion) or creversion == 'xx.7':
                reverison.send_keys("A")
            else:
                reverison.send_keys(creversion)
        branditem = driver.find_element(By.ID, "BrndMrkItm")
        Select(branditem).select_by_value("No marking")
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
        finish_button = driver.find_element(By.ID, "ext-gen39")
        finish_button.click()
        try:
            WebDriverWait(driver, 5).until(EC.alert_is_present())
        except NoAlertPresentException:
            driver.switch_to.window(home)
            wait.until(EC.visibility_of_element_located((By.ID, 'psbIFrame')))
            #     # Switch to the frame //Switch to base frame  switchTo().defaultContent()
            driver.switch_to.frame('psbIFrame')
            # edit_usage_page = driver.find_elements(By.XPATH, "//span[contains(text(),'Edit Usage Attributes')]")
            qty_input_new = wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                         "//label[contains(text(),'*Quantity')]/../../following-sibling::td[1]//input")))  # complex locate element xpath
            qty_input_new.clear()
            qty_input_new.send_keys(cqty)
            time.sleep(1.5)
            tech_new = wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                    "//label[contains(text(),'*Technical')]/../../following-sibling::td[1]//input")))
            tech_new.clear()
            tech_new.send_keys(ctechpart)
            time.sleep(1.5)
            driver.find_element(By.XPATH, "//button[contains(text(),'OK')]").click()
        except NoSuchWindowException:
            driver.switch_to.window(home)
            wait.until(EC.visibility_of_element_located((By.ID, 'psbIFrame')))
            #     # Switch to the frame //Switch to base frame  switchTo().defaultContent()
            driver.switch_to.frame('psbIFrame')
            # edit_usage_page = driver.find_elements(By.XPATH, "//span[contains(text(),'Edit Usage Attributes')]")
            qty_input_new = wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                         "//label[contains(text(),'*Quantity')]/../../following-sibling::td[1]//input")))  # complex locate element xpath
            qty_input_new.clear()
            qty_input_new.send_keys(cqty)
            time.sleep(1.5)
            tech_new = wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                    "//label[contains(text(),'*Technical')]/../../following-sibling::td[1]//input")))
            tech_new.clear()
            tech_new.send_keys(ctechpart)
            time.sleep(1.5)
            driver.find_element(By.XPATH, "//button[contains(text(),'OK')]").click()
        else:
            driver.switch_to.alert.accept()
            driver.find_element(By.ID, 'ext-gen41').click()
            driver.switch_to.window(home)
            wait.until(EC.visibility_of_element_located((By.ID, 'psbIFrame')))
            #     # Switch to the frame //Switch to base frame  switchTo().defaultContent()
            driver.switch_to.frame('psbIFrame')
            existpage(driver, wait, parentnumber, cnumber, cqty, ctechpart)
    else:
        return

    # for insert exist items


def existpage(driver, wait, parentnumbere, number, eqty, etechpart):
    try:
        click_parent(driver, parentnumbere)
    except NoSuchElementException:
        js = 'document.getElementsByClassName("x-grid3-scroller")[0].scrollTop=300'
        driver.execute_script(js)
        # Can not use element_to_be_clickable, have to use visibility_of_element_located
    wait.until(EC.visibility_of_element_located(
        (By.XPATH, "//*[@class='x-grid3-scroller']//span[contains(text(),'" + parentnumbere + "')]")))

    # #enable_state = insert_exist_button.get_attribute("aria-disabled")    get_attribute does not work, do not know why.
    # enable_state = insert_exist_button.is_enabled()
    # print(str(enable_state))
    # if enable_state == False:
    #     return
    # insert_new_button = driver.find_element(By.XPATH, "//button[contains(text(),'Insert New')]")
    # new_button_state = insert_new_button.get_attribute("disabled")
    # print(str(new_button_state))
    try:
        WebDriverWait(driver, 4).until(
            EC.element_attribute_to_include((By.XPATH, "//button[contains(text(),'Insert New')]"), 'disabled'))
    except TimeoutException:
        insert_exist_button = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Insert Existing')]")))
        try:
            insert_exist_button.click()
        except ElementClickInterceptedException:
            ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'OK')]")))
            ok_button.click()
        # existframe = wait.until(EC.visibility_of_element_located((By.XPATH, "//*[@id='ext-gen1']/iframe[2]")))
        # existframe = driver.find_element(By.XPATH, "//*[@id='ext-gen1']/iframe[2]")
        # driver.switch_to.frame(existframe)
        numberex = wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "//label[contains(text(),'Number:')]/following-sibling::div//input")))
        # numberex = driver.find_element(By.XPATH, "//label[contains(text(),'Number:')]/following-sibling::div//input")
        numberex.send_keys(number)
        time.sleep(1)
        searchbutton = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Search')]")))
        searchbutton.click()
        time.sleep(1)
        try:
            ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'OK')]")))
            # okbutton = driver.find_element(By.XPATH, "//button[contains(text(),'OK')]")
            ok_button.click()
        except ElementClickInterceptedException:
            searchbutton.click()
            time.sleep(0.5)
            ok_button.click()
        # action_chains = ActionChains(driver)
        # action_chains.double_click(ok_button).perform()
        time.sleep(2)
        add_qty_page = driver.find_elements(By.XPATH, "//span[contains(text(),'Edit Usage Attributes')]")
        if len(add_qty_page) == 1:
            qty_input = wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                     "//label[contains(text(),'*Quantity')]/../../following-sibling::td[1]//input")))  # loacte two same input fields
            qty_input.clear()
            qty_input.send_keys(eqty)
            time.sleep(1.5)
            tech = wait.until(EC.visibility_of_element_located((By.XPATH,
                                                                "//label[contains(text(),'*Technical')]/../../following-sibling::td[1]//input")))
            tech.clear()
            tech.send_keys(etechpart)
            time.sleep(1.5)
            driver.find_element(By.XPATH, "//button[contains(text(),'OK')]").click()
        time.sleep(2)
        confirm_no = driver.find_elements(By.XPATH, "//button[contains(text(),'No')]")
        if len(confirm_no) == 1:
            confirm_no_button = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'No')]")))
            confirm_no_button.click()
        # return False
        # Edit Usage Attributes page
        # editusageframe = wait.until(
        #     EC.visibility_of_element_located((By.XPATH, "//iframe[@class='ext-shim x-ignore']")))
        # driver.switch_to.frame(editusageframe)
        #     two_inputs = driver.find_elements(By.XPATH, "//input[@role='spinbutton']")  # loacte two same input fields
        #     two_inputs[1].send_keys(cqty)
        #     driver.find_element(By.XPATH, "//*[@role='alert']/preceding-sibling::div//input").send_keys(
        #         etechpart)
        #     driver.find_element(By.XPATH, "//button[contains(text(),'OK')]").click()
        # except Exception:
        #     pass
        # wait.until(
        #     EC.visibility_of_element_located((By.ID, 'psbIFrame'))
        # )
        # Switch to the frame //Switch to base frame  switchTo().defaultContent()
        # //*[@class="x-tree3-node-text"] to locate the items in the Identity
        # driver.switch_to.frame("psbIFrame")
    else:
        return


def getfile():
    path = sys.argv[1]
    return path


def all_level(driver):
    icon_xpath = "//*[@style='width:16px;height:16px;background:url(https://pdmlink.niladv.org/Windchill/gwt/com.ptc.windchill.wncgwt.WncGWT/73C4E3626360BA866F88BA514C3354A8.cache.png) no-repeat -66px 0px;']"
    trangle_icons = driver.find_elements(By.XPATH, icon_xpath)
    while len(trangle_icons) > 0:
        for icon in trangle_icons:
            try:
                WebDriverWait(driver, 10).until(
                    EC.visibility_of_all_elements_located((By.XPATH, icon_xpath)))
                driver.execute_script("arguments[0].click();", icon)

            except StaleElementReferenceException:
                trangle_icons = driver.find_elements(By.XPATH, icon_xpath)
            except TimeoutException:
                trangle_icons = driver.find_elements(By.XPATH, icon_xpath)
            if len(trangle_icons) == 0:
                break

        trangle_icons = driver.find_elements(By.XPATH, icon_xpath)


def main():
    # flag = False
    mdriver, mwait, mhome = initdriver()
    filepath = getfile()
    with open(filepath, 'rb') as f:
        data = pd.read_excel(f, dtype=str)
        # print(data.columns)
        # print(data['Number'][0], data['Name'][0])

        item = mwait.until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, data['Name'][0]))
        )
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
        # action_chain = ActionChains(mdriver)
        all_level(mdriver)
        for i in range(1, len(data['Level'])):
            # print(i, data['Level'][i])
            if int(data['Level'][i]) == int(data['Level'][i - 1]) + 1:
                # print(i, data['Level'][i])
                if data['Existing\n/New'][i] == 'Yes':
                    existpage(mdriver, mwait, data['Number'][i - 1], data['Number'][i], data['Qty'][i],
                              data['Technical Part'][i])
                # run the insert exist page
                elif data['Existing\n/New'][i] == 'No':
                    newpage(mdriver, mwait, mhome, data['Number'][i - 1], data['Number'][i], data['Name'][i],
                            data['Assembly Mode'][i].lower(),
                            data['Source'][i].lower(),
                            data['Revision'][i], data['Accessory'][i],
                            data['Spare\n Part'][i],
                            data['Gathering Part'][i],
                            data['Critical Characteristic'][i], data['Qty'][i], data['Technical Part'][i])
            # run the insert new page i是当前行，i-1是父行
            elif int(data['Level'][i]) == int(data['Level'][i - 1]):
                # print(i, "dddd")
                for j in range(i - 1, -1, -1):
                    if int(data['Level'][j]) == int(data['Level'][i]) - 1:
                        # print(j, data['Number'][j])
                        if data['Existing\n/New'][i] == 'Yes':
                            existpage(mdriver, mwait, data['Number'][j], data['Number'][i], data['Qty'][i],
                                      data['Technical Part'][i])
                        # run the insert exist page
                        elif data['Existing\n/New'][i] == 'No':
                            # run the insert new page i是当前行，j是父行
                            newpage(mdriver, mwait, mhome, data['Number'][j], data['Number'][i], data['Name'][i],
                                    data['Assembly Mode'][i].lower(),
                                    data['Source'][i].lower(),
                                    data['Revision'][i], data['Accessory'][i],
                                    data['Spare\n Part'][i],
                                    data['Gathering Part'][i],
                                    data['Critical Characteristic'][i], data['Qty'][i],
                                    data['Technical Part'][i])
                        break
            elif int(data['Level'][i]) < int(data['Level'][i - 1]):
                # print(i, "aaaa")
                # for jj in range(len(data['Level'][1:i]), -1, -1):
                for jj in range(i - 1, -1, -1):
                    if int(data['Level'][jj]) == int(data['Level'][i]) - 1:
                        # print(jj, data['Number'][jj])
                        if data['Existing\n/New'][i] == 'Yes':
                            # run the insert exist page
                            existpage(mdriver, mwait, data['Number'][jj], data['Number'][i], data['Qty'][i],
                                      data['Technical Part'][i])
                        elif data['Existing\n/New'][i] == 'No':
                            # run the insert new page, i是当前行，jj是父行
                            newpage(mdriver, mwait, mhome, data['Number'][jj], data['Number'][i],
                                    data['Name'][i],
                                    data['Assembly Mode'][i].lower(),
                                    data['Source'][i].lower(),
                                    data['Revision'][i], data['Accessory'][i],
                                    data['Spare\n Part'][i],
                                    data['Gathering Part'][i],
                                    data['Critical Characteristic'][i], data['Qty'][i],
                                    data['Technical Part'][i])
                        break


if __name__ == "__main__":
    main()

    # number=WebDriverWait(driver, 20).until(
    #     EC.element_to_be_clickable((By.XPATH, "//*[@id='number']"))
    # )
    # https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1
    # Part https://pdmlink.niladv.org/Windchill/app/#ptc1/tcomp/infoPage?ContainerOid=OR%3Awt.pdmlink.PDMLinkProduct%3A767792377&oid=OR%3Awt.folder.SubFolder%3A767792837&u8=1
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
