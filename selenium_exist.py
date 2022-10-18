import sys
import time

import pandas as pd
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC


def initdriver():
    s = Service(r'C:\sources\edgedriver_win64\msedgedriver.exe')
    wdriver = webdriver.Edge(service=s)
    # wdriver.maximize_window()
    wdriver.set_window_size(1925, 1085)
    url = "https://pdmlink.niladv.org/Windchill/app/"
    wdriver.get(url)
    wwait = WebDriverWait(wdriver, 60)
    return wdriver, wwait


def getfile():
    path = sys.argv[1]
    return path


def main():
    driver, wait = initdriver()
    search_input = wait.until(EC.element_to_be_clickable((By.ID, "gloabalSearchField")))
    file_path = getfile()
    with open(file_path, 'r+b') as f:
        data = pd.read_excel(f, dtype=str)
        # print(data.head())
        for i in range(0, len(data['Number'])):
            search_input.send_keys(data['Number'][i])
            search_input.send_keys(Keys.ENTER)
            # element_to_be_clickable does not work, so change to visibility_of_element_located
            wait.until(
                EC.visibility_of_element_located((By.XPATH,
                                                  "//*[@id='wt.fc.Persistable.defaultSearchViewfilterSelect']")))  # //a[contains(text(),'56508734')]/../../preceding-sibling::td[1]//img   小螺丝图片
            time.sleep(2)
            screw_icon = driver.find_elements(By.XPATH, "//a[contains(text(),'" + data['Number'][
                i] + "')]/../../preceding-sibling::td[1]//img")
            release_states = driver.find_elements(By.XPATH, "//a[contains(text(),'" + data['Number'][
                i] + "')]/../../following-sibling::td[6]//div[@class='x-grid3-cell-inner x-grid3-col-state']")
            if len(screw_icon) > 0:
                for state in release_states:
                    text = wait.until(
                        EC.visibility_of_element_located((By.XPATH, "//a[contains(text(),'" + data['Number'][
                            i] + "')]/../../following-sibling::td[6]//div[@class='x-grid3-cell-inner x-grid3-col-state']"))).text
                    data["Existing\n/New"][i] = text
            else:
                data["Existing\n/New"][i] = 'New'
        data.to_excel(file_path, sheet_name='Sheet1', index=False, header=True)


if __name__ == "__main__":
    main()
    print("Process completed!")
