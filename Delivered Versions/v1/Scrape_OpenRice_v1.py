from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
import time
import os
from datetime import datetime
import pandas as pd
import warnings
import re
import sys
import xlsxwriter
import platform
from multiprocessing import freeze_support
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    #chrome_options  = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.page_load_strategy = 'eager'
    # disable location prompts & disable images loading
    prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)

    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    driver = uc.Chrome(version_main = ver, options=chrome_options)
    
    driver.set_window_size(1920, 1080, driver.window_handles[0])
    driver.maximize_window()
    driver.set_page_load_timeout(300)

    return driver

def scrape_restaurants(driver, output1, output2, page):

    print('-'*75)
    if 'new-restaurants' in page:
        print('Scraping The New Restaurants Links ...')
        res_type = 'New'
    elif 'best-rating' in page:
        print('Scraping The Best Rating Restaurants Links ...')
        res_type = 'Best Rating'
    elif 'most-popular' in page:
        print('Scraping The Most Popular Restaurants Links ...')
        res_type = 'Most Popular'
    elif 'most-bookmarked' in page:
        print('Scraping The Most Bookmarked Restaurants Links ...')
        res_type = 'Most Bookmarked'
    else:
        print('Scraping The Best Dessert Restaurants Links ...')
        res_type = 'Best Dessert'

    print('-'*75)
    driver.get(page)
    time.sleep(3)

    # processing the lazy loading of the restaurants
    while True:  
        try:
            height1 = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {height1})")
            time.sleep(5)
            height2 = driver.execute_script("return document.body.scrollHeight")
            if int(height2) == int(height1):
                break
        except Exception as err:
            break

    # getting the full restaurants list
    links = []
    # scraping restaurants urls
    if 'new-restaurants' in page:
        selector =  "a.poi-list-cell-info-title"
    else:
        selector = 'a.chart-poi-name'
    restaurants = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector)))
    nres = 0
    for res in restaurants:
        try:
            nres += 1
            print(f'Scraping the url for restaurant {nres}')
            link = res.get_attribute('href')
            links.append(link)
        except:
            pass

    # scraping restaurants details
    print('-'*75)
    print('Scraping Restaurants Details...')
    print('-'*75)
    n = len(links)
    data = pd.DataFrame()
    reviews = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            driver.get(link)           
            details, review = {}, {}
            print(f'Scraping the details of restaurant {i+1}\{n}')

            # Restaurant name Chinese and English
            name_en, name_ch = '', ''              
            try:
                name_ch = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.name"))).get_attribute('textContent').strip()
                name_en = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.smaller-font-name"))).get_attribute('textContent').strip()
                # check if the chinese name is English 
                asian = re.findall(r'[\u3131-\ucb4c]+',name_ch)
                # English name is found in the chinese name
                if not asian and name_en == '':
                    name_en = name_ch
                    name_ch = ''
                # English and Chinese names are in one field
                elif asian and name_en == '':
                    name = ''.join(asian)
                    name_en = name_ch.replace(name, '')
                    name_ch = name 

                ascii_ = re.findall(r'[ -~]',name_en)
                # English and Chinese names are in one field
                if ascii_ and name_ch == '':
                    name = ''.join(ascii_)
                    name_ch = name_en.replace(name, '')
                    name_en = name    
                # Chinese name is found in the English name
                elif ascii_ and name_ch != '':
                    name_en = ''.join(ascii_)
            except:
                print(f'Warning: failed to scrape the name for restaurant: {link}')               
                
            details['Name_Chinese'] = name_ch
            details['Name_English'] = name_en.replace('(' ,'').replace(')', '').strip()
                                    
            # Price range 
            price = ''
            try:
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='header-poi-price dot-separator']")))
                price = wait(div, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('textContent')
            except:
                pass
                    
            details['Price_Range'] = price            
             
            # Restaurant category 
            cat = ''
            try:
                cat = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='header-poi-categories dot-separator']"))).get_attribute('textContent').replace('\n', '').strip().replace('                        ', ', ')
            except:
                pass          
                
            details['Category'] = cat            
            
            # Address
            add = ''
            try:
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='address-info-section']")))
                add = wait(div, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.content"))).get_attribute('textContent').strip()
            except:
                pass          
                
            details['Address'] = add           
                               
            # Region
            region = ''
            try:
                region = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='header-poi-district dot-separator']"))).get_attribute('textContent').strip()
            except:
                pass          
                
            details['Region'] = region 
            
            # Telephone number
            tel = ''
            try:
                sec = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "section[class='telephone-section']")))
                tags = wait(sec, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.content")))
                for tag in tags:
                    tel += tag.get_attribute('textContent').strip() + ', '
                tel = tel[:-2]
            except:
                pass          
                
            details['Phone'] = tel                                    
            # Introduction
            intro = ''
            try:
                sec = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "section[class='introduction-section']")))
                intro = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.content"))).get_attribute('textContent').replace('.. ', '').replace('\n', '').strip()
            except:
                pass          
                
            details['Introduction'] = intro              
            
            # number of faces
            smily, ok, sad = '', '', ''
            try:
                div = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='header-smile-section']")))
                faces = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.score-div")))
                smily = faces[0].get_attribute('textContent').strip()
                ok = faces[1].get_attribute('textContent').strip()
                sad = faces[2].get_attribute('textContent').strip()
            except:
                pass          
                
            details['Face_Smiley'] = smily              
            details['Face_Ok'] = ok              
            details['Face_Sad'] = sad             
            
            # Review rating
            rating = ''
            try:
                rating = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='header-score']"))).get_attribute('textContent').strip()
            except:
                pass          
                
            details['Rating'] = rating               
            
            # Number of bookmarks
            books = ''
            try:
                books = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='header-bookmark-count js-header-bookmark-count']"))).get_attribute('textContent').strip()
            except:
                pass          
                
            details['Bookmarks'] = books                                
            details['Openrice_Link'] = link 
            if 'new-restaurants' in page:
                details['Rank'] = ''
            else:
                details['Rank'] = i+1

            details['Restaurant_Type'] = res_type
            ## scraping restaurants reviews
            try:
                url = link + '/reviews'
                driver.get(url)
                while True:
                    sections = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "section[class='sr2-review-list2-main-content-section']")))
                    for sec in sections:
                        try:
                            title = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.review-title"))).get_attribute('textContent').strip()
                            date = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[itemprop='datepublished']"))).get_attribute('textContent').strip()
                            nviews = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[class='view-count-container']"))).get_attribute('textContent').split()[0].strip()
                            des = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[itemprop='description']"))).get_attribute('textContent').strip()
                            if name_ch != '':
                                review['Restaurant_Name'] = name_ch
                            else:
                                review['Restaurant_Name'] = name_en
                            review['Review_Title'] = title
                            review['Review_Date'] = date
                            review['Review_Views'] = nviews
                            review['Review_Content'] = des

                            reviews = reviews.append([review.copy()])
                        except:
                            pass 
                    # moving to the next page
                    try:
                        a = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='pagination-button next js-next']")))
                        url = a.get_attribute('href')
                        driver.get(url)
                    except:
                        break
            except:
                pass

            # appending the output to the datafame       
            data = data.append([details.copy()])
            # saving data to csv file each 100 links
            #if np.mod(i+1, 50) == 0:
            #    print('Outputting scraped data ...')
            #    data.to_excel(output1, index=False)
            #    reviews.to_excel(output2, index=False)

        except Exception as err:
            pass
            #print(str(err))
           
    # output to excel
    df1 = pd.read_excel(output1)
    df2 = pd.read_excel(output2)
    df1 = df1.append(data)
    df2 = df2.append(reviews)
    df1.to_excel(output1, index=False)
    df2.to_excel(output2, index=False)
 

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\scraped_data\\' + stamp
    if os.path.exists(path):
        os.remove(path) 
    os.makedirs(path)

    file1 = f'OpenRice_{stamp}.xlsx'
    file2 = f'OpenRice_Comments_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
        output2 = path.replace('\\', '/') + "/" + file2
    else:
        output1 = path + "\\" + file1
        output2 = path + "\\" + file2   

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    worksheet = workbook1.add_worksheet()
    workbook1.close()    
    workbook2 = xlsxwriter.Workbook(output2)
    worksheet = workbook2.add_worksheet()
    workbook2.close()

    return output1, output2

def main():
    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    output1, output2 = initialize_output()
    homepages = ["https://www.openrice.com/en/hongkong/chart/best-rating", "https://www.openrice.com/en/hongkong/chart/most-popular", "https://www.openrice.com/en/hongkong/chart/most-bookmarked", "https://www.openrice.com/en/hongkong/chart/best-dessert", "https://www.openrice.com/en/hongkong/new-restaurants"]
    
    try:
        driver = initialize_bot()
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()
    for page in homepages:
        try:
            scrape_restaurants(driver, output1, output2, page)
        except Exception as err: 
            #print(f'Error:\n')
            #print(str(err))
            driver.quit()
            time.sleep(5)
            driver = initialize_bot()

    driver.quit()
    print('-'*100)
    elapsed_time = round(((time.time() - start)/60), 2)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')

if __name__ == '__main__':

    main()