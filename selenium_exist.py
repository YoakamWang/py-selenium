import sys

import pandas as pd
from selenium import webdriver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC


def initdriver():
    s = Service(r'C:\sources\edgedriver_win64\msedgedriver.exe')
    # driver = webdriver.Edge(executable_path=r'C:\sources\edgedriver_win64\msedgedriver.exe')
    # driver.get("https://www.baidu.com")
    wdriver = webdriver.Edge(service=s)
    wdriver.maximize_window()
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
        for i in range(1, len(data['Number'])):
            search_input.send_keys(data['Number'][i])
            search_input.send_keys(Keys.ENTER)
            # // *[ @ id = "wt.fc.Persistable.defaultSearchViewfilterSelect"]
            wait.until(
                EC.visibility_of_element_located((By.XPATH,
                                                  "//*[@id='wt.fc.Persistable.defaultSearchViewfilterSelect']")))  # //a[contains(text(),'56508734')]/../../preceding-sibling::td[1]//img   小螺丝图片

            screw_icon = driver.find_elements(By.XPATH, "//a[contains(text(),'" + data['Number'][
                i] + "')]/../../preceding-sibling::td[1]//img")
            if len(screw_icon) > 0:
                data["Existing\n/New"][i] = 'Yes'
            else:
                data["Existing\n/New"][i] = 'No'
        data.to_excel(file_path, sheet_name='Sheet1', index=False, header=True)


if __name__ == "__main__":
    main()
