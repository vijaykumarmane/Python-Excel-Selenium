from selenium import webdriver
import os
import xlwings as xw
import time
import re

step = str(input("Step to Run (1 or 2 and press Enter):"))

if step == '2':
    print("Running Website seach")

    # Open saved workbook
    wb = xw.Book(os.path.expanduser("~") +
                 "\\Desktop\\Contact Details.xlsx")
    sht = wb.sheets[0]

    driver = webdriver.Chrome(os.path.expanduser(
        '~') + "\\Desktop\\chromedriver.exe")

    driver.set_page_load_timeout(15)

    def fetchInfo(link):
        email, phone = [], []

        try:
            driver.get(link)
        except:
            pass
        try:
            all_visible_text = driver.find_element_by_tag_name('body').text
        except:
            return [], []
        email = re.findall(
            r"[a-zA-Z0-9_.+-]+@[a-zA-Z-]+\.[a-zA-Z.]+", all_visible_text)
        if len(email) == 0:
            email = re.findall(
                r"[a-zA-Z0-9_.+-]+@[a-zA-Z-]+\.[a-zA-Z.]+\.[a-zA-Z.]", all_visible_text)
        # phone = re.findall(r"\+[-()\s\d]+?(?=\s*[+<])", all_visible_text)
        all_digit_stram = re.findall('([()\+0-9 -]+)\n', all_visible_text)

        for no in all_digit_stram:
            no = str(no)
            no = no.replace("-", " ")
            no = no.strip()
            
            if len(no) == 0:
                continue
            if "12345" not in no:
                if "---" not in no:
                    if "-" != no[0]:
                        if len(no) > 10:
                            phone = str(no)
                            break
        if "bloomberg" in link or "yello" in link or "zoominfo" in link:
            print(email, phone)
            return email, phone
            
        if len(email) == 0:
            try:
                contact_link = driver.find_element_by_partial_link_text(
                    "Contact").click()
            except:
                try:
                    contact_link = driver.find_element_by_partial_link_text(
                        "CONTACT").click()
                except:
                    try:
                        contact_link = driver.find_element_by_partial_link_text(
                            "touch").click()
                    except:
                        try:
                            contact_link = driver.find_element_by_partial_link_text(
                                "Touch").click()
                        except:
                            pass
            try:
                all_visible_text = driver.find_element_by_tag_name('body').text
            except:
                return [], []
            email = re.findall(
                r"[a-zA-Z0-9_.+-]+@[a-zA-Z-]+\.[a-zA-Z.]+", all_visible_text)
            if len(email) == 0:
                email = re.findall(
                    r"[a-zA-Z0-9_.+-]+@[a-zA-Z-]+\.[a-zA-Z.]+\.[a-zA-Z.]+", all_visible_text)
            # phone = re.findall(r"\+[-()\s\d]+?(?=\s*[+<])", all_visible_text)
            all_digit_stram = re.findall('([()\+0-9 -]+)\n', all_visible_text)
            for no in all_digit_stram:
                no = str(no)
                no = no.replace("-", " ")
                no = no.strip()
                
                if len(no) == 0:
                    continue
                if "12345" not in no:
                    if "---" not in no:
                        if "-" != no[0]:
                            if len(no) > 10:
                                phone = str(no)
                                break
        print(email, phone)
        return email, phone

    r = 1
    count = 1
    i = 2
    while sht.range("B" + str(i)).value is not None:
        if sht.range("E" + str(i)).value == "Googled":
            companyName = sht.range("B" + str(i)).value
            prevName = sht.range("B" + str(i - 1)).value
            print(i, end=" ")
            if companyName != prevName:
                for col in ["F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U"]:
                    if sht.range(col + str(i)).value is None:
                        continue

                    if sht.range(col + str(i)).api.Font.Bold:
                        link = sht.range(col + str(i)).value
                        
                        link = link.split("\n")[1]
                        output = fetchInfo(link)
                        r += 1
                        sht.range("X" + str(i)).value = ";".join(output[0])
                        if len(output[0]) != 0:
                            for m in output[0]:
                                m = m.lower().strip()
                                if ".png" not in m:
                                    if ".gif" not in m:
                                        if "panjiva" not in m:
                                            if "domain" not in m:
                                                if "example" not in m:
                                                    if "name" not in m:
                                                        sht.range(
                                                            "C" + str(i)).value = m
                                                        count += 1
                                                        break
                        sht.range("Y" + str(i)).value = str(output[1])
                        if len(output[1]) != 0 and sht.range("D" + str(i)).value is None:
                            sht.range("D" + str(i)).value = str(output[1])
                        sht.range("E" + str(i)).value = "Site Searched"
                        break
        i += 1
        #break
    print("Done", count, " out of", i)
    driver.close()

if step == '1':
    print("Running google seach results")

    wb = xw.Book(os.path.expanduser("~") +
                 "\\Desktop\\Contact Details.xlsx")
    sht = wb.sheets[0]

    driver = webdriver.Chrome(os.path.expanduser(
        '~') + "\\Desktop\\chromedriver.exe")


    def duckduck(companyName):

        url = "https://duckduckgo.com/?q=" + \
            companyName.replace("& ", "")
        url = url.replace(' ', '+')
        driver.get(url)
        elems = driver.find_elements_by_tag_name("h2")
        href_links = []

        for i in elems:
            try:
                link = i.text + "\n" + \
                    i.find_elements_by_tag_name('a')[0].get_attribute("href")
            except:
                break

            if "linked" in link:
                continue
            if "panjiva" in link:
                continue
            if "facebook" in link:
                continue

            href_links.append(link)

        return href_links


    def ecosia(companyName):

        url = "https://www.ecosia.org/search?q=" + \
            companyName.replace("& ", "")
        url = url.replace(' ', '+')
        driver.get(url)
        elems = driver.find_elements_by_tag_name("h2")

        href_links = []

        for i in elems:
            try:
                link = i.text + "\n" + \
                    i.find_elements_by_tag_name('a')[0].get_attribute("href")
            except:
                continue

            if "linked" in link:
                continue
            if "panjiva" in link:
                continue
            if "facebook" in link:
                continue
            if "merriam" in link:
                continue

            href_links.append(link)

        return href_links


    def googleSearchTile(driver, phone):
        '''
        url = "https://www.google.com/search?q=" + \
            companyName.replace("& ", "") + " Contact Us"
        url = url.replace(' ', '+')
        driver.get(url)
        '''

        all_visible_text = driver.find_element_by_tag_name('body').text
        if len(str(phone)) < 4:
            try:
                phone = re.search("Phone: ([+0-9- ]+)", all_visible_text).group(1)
            except:
                phone = ""

        email = re.findall(
            r"[a-zA-Z0-9_.+-]+@[a-zA-Z-]+\.[a-zA-Z.]+", all_visible_text)

        if len(email) == 0:
            email = re.findall(
                r"[a-zA-Z0-9_.+-]+@[a-zA-Z-]+\.[a-zA-Z.]+\.[a-zA-Z]+", all_visible_text)

        website = ""

        for ele in driver.find_elements_by_xpath("//a[@role='button']"):
            if "Website" == ele.text:
                website = ele.get_attribute("href")
        try:
            title = driver.find_elements_by_xpath(
                "//h2[@data-attrid='title']")[0].text
        except:
            title = ""

        emailStr = ""

        if len(email) != 0:
            emailStr = ";".join(email)

        if len(title) == 0:
            title = ""
        if len(website) == 0:
            website = ""

        tile = title + "\n" + website + "\n" + emailStr + "\n" + phone

        return tile, phone


    def google(companyName):

        url = "https://www.google.com/search?q=" + \
            companyName.replace("& ", "")  # + " Contact us"
        url = url.replace(' ', '+')
        driver.get(url)
        elems = driver.find_elements_by_tag_name("a")

        href_links = []

        for i in elems:
            try:
                link = i.find_elements_by_tag_name(
                    'h3')[0].text + "\n" + i.get_attribute("href")
            except:
                continue

            if "linked" in link:
                continue
            if "panjiva" in link:
                continue
            if "facebook" in link:
                continue
            if "merriam" in link:
                continue

            href_links.append(link)

        tile, phone = googleSearchTile(driver, "0")

        iffound = tile.split("\n")[2]
        '''
        if len(iffound) < 4:
            url = "https://www.google.com/search?q=" + \
                companyName.replace("& ", "") + " email"
            driver.get(url)
            tile, phone = googleSearchTile(driver, phone)
            iffound = tile.split("\n")[2]

        if len(iffound) < 4:
            url = "https://www.google.com/search?q=" + \
                companyName.replace("& ", "") + " Contact Us"
            driver.get(url)
            tile, phone = googleSearchTile(driver, phone)
    '''
        return tile, href_links


    found = 0
    i = 2
    while sht.range("B" + str(i)).value is not None:
        if i % 3 == 0:
            driver.close()
            driver = webdriver.Chrome(os.path.expanduser(
                '~') + "\\Desktop\\chromedriver.exe")

        if sht.range("C" + str(i)).value is None and sht.range("G" + str(i)).value is None:
            companyName = sht.range("B" + str(i)).value
            prevName = sht.range("B" + str(i - 1)).value
            if companyName != prevName:

                output = google(companyName)

                sht.range("F" + str(i)).value = output[0]

                got = output[0].split("\n")
                if sht.range("C" + str(i)).value is None:
                    sht.range("C" + str(i)).value = got[2]
                if sht.range("D" + str(i)).value is None:
                    sht.range("D" + str(i)).value = got[3]
                print(i, found, got[2:])
                found += 1
                sht.range("G" + str(i)).value = output[1]

                sht.range("E" + str(i)).value = "Googled"

        i += 1
        # break
    wb.save()
    driver.close()

    print('Done')
else:
    print("Hi Hi Hi, Wrong input :;")
    time.sleep(5)
