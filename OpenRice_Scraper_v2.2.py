from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
import time
import os
from datetime import datetime, timedelta
import pandas as pd
import warnings
import re
import sys
import csv
import xlsxwriter
from multiprocessing import freeze_support
warnings.filterwarnings('ignore')

def initialize_bot():

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--headless=new')
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.page_load_strategy = 'normal'
    # disable location prompts & disable images loading
    prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080, driver.window_handles[0])
    driver.maximize_window()
    driver.set_page_load_timeout(20000)

    return driver

def scrape_restaurants(driver, output1, output2, page, settings):

    stamp = datetime.now().strftime("%d/%m/%Y")
    if isinstance(page, dict):
        print('-'*75)
        res_type = list(page.keys())[0]
        print(f'Scraping The {res_type} Restaurants ...')
        print('-'*75)      
    elif 'new-restaurants' in page:
        if settings["New Restaurants"] == 0: return
        print('-'*75)
        print('Scraping The New Restaurants ...')
        res_type = 'New'
    elif 'best-rating' in page:
        if settings["Best Rated Restaurants"] == 0: return
        print('-'*75)
        print('Scraping The Best Rating Restaurants ...')
        res_type = 'Best Rating'
    elif 'most-popular' in page:
        if settings["Most Popular Restaurants"] == 0: return
        print('-'*75)
        print('Scraping The Most Popular Restaurants ...')
        res_type = 'Most Popular'
    elif 'most-bookmarked' in page:
        if settings["Most Bookmarked Restaurants"] == 0: return
        print('-'*75)
        print('Scraping The Most Bookmarked Restaurants ...')
        res_type = 'Most Bookmarked'
    elif 'best-dessert' in page:
        if settings["Best Dessert Restaurants"] == 0: return
        print('-'*75)
        print('Scraping The Best Dessert Restaurants ...')
        res_type = 'Best Dessert'

    links = []
    if isinstance(page, dict):
        links = page[res_type]
    else:
        print('-'*75)
        driver.get(page)
        time.sleep(3)

        res_limit = settings['Restaurants Limit']
        # scraping restaurants urls
        if 'new-restaurants' in page:
            selector =  "a.poi-list-cell-info-title"
        else:
            selector = 'a.chart-poi-name'
        # processing the lazy loading of the restaurants
        while True:  
            try:
                height1 = driver.execute_script("return document.body.scrollHeight")
                driver.execute_script(f"window.scrollTo(0, {height1})")
                time.sleep(5)
                height2 = driver.execute_script("return document.body.scrollHeight")
                restaurants = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector)))
                if len(restaurants) >= res_limit and res_limit > 0: 
                    break
                if int(height2) == int(height1):
                    break
            except Exception as err:
                break

        # getting the full restaurants list      
        restaurants = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, selector)))
        # applying the limit
        if res_limit > 0:
            restaurants = restaurants[:res_limit]
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
        #print('-'*75)
        print('Scraping Restaurants Details')
        #print('-'*75)

    n = len(links)
    data = pd.DataFrame()
    reviews = pd.DataFrame()
    for i, link in enumerate(links):
        search_name, search_loc = ' ', ' '
        try:
            if res_type == 'Search Result':
                search_name = link[1]
                search_loc = link[2]
                link = link[0]

            driver.get(link)           
            details, review = {}, {}
            print(f'Scraping the details of restaurant {i+1}\{n}')

            # Restaurant name Chinese and English
            name_en, name_ch = '', ''              
            try:
                name_ch = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span.name"))).get_attribute('textContent').strip()
                name_en = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.smaller-font-name"))).get_attribute('textContent').strip()

                ## check if the chinese name is English 
                #asian = re.findall(r'[\u3131-\ucb4c]+',name_ch)
                #asian2 = re.findall(r'[\u3131-\ucb4c]+',name_en)
                ## English name is found in the chinese name
                #if not asian and name_en == '':
                #    name_en = name_ch
                #    name_ch = ''
                ## English and Chinese names are in one field
                #elif asian and name_en == '':
                #    name = ''.join(asian)
                #    name_en = name_ch.replace(name, '')
                #    name_ch = name                 
                                   
                ## English and Chinese names are in one field
                #elif asian2 and name_ch == '':
                #    name = ''.join(asian2)
                #    name_en = name_en.replace(name, '')
                #    name_ch = name 
                ################################################################
                if name_en == '' and name_ch != '':
                    name_en = name_ch
                elif name_ch == '' and name_en != '':
                    name_ch = name_en
            except:
                print(f'Warning: failed to scrape the name for restaurant: {link}')               
                
            details['Name_Chinese'] = name_ch
            details['Name_English'] = name_en
                                    
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
                intro = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.content"))).get_attribute('textContent').replace('.. ', '').replace('\n', '').replace('continue reading', '').strip()
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

            if res_type == 'New' or res_type == 'User Input' or res_type == 'Search Result' or 'Search Link' in res_type:
                details['Rank'] = ''
            else:
                details['Rank'] = i+1

            if res_type == 'Search Result':
                details['Restaurant_Type'] = res_type + f'-{search_name}, {search_loc}'
            else:
                details['Restaurant_Type'] = res_type

            details['Extraction Date'] = stamp
            ## scraping restaurants reviews
            if settings["Scrape Reviews"] != 0:
                rev_limit = settings["Reviews Limit"]
                try:
                    url = link + '/reviews'
                    driver.get(url)
                    end = False
                    nrevs = 0
                    while True:
                        sections = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[itemprop='review']")))
                        for sec in sections:
                            try:
                                title = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.review-title"))).get_attribute('textContent').strip()
                                rev_link = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.title"))).get_attribute('href')
                                date = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[itemprop='datepublished']"))).get_attribute('textContent').strip()
                                if 'day(s) ago' in date:
                                    try:
                                        num = int(date.replace('day(s) ago', '').strip())
                                        date = datetime.today() - timedelta(days=num)
                                        date = date.strftime('%Y-%m-%d')
                                    except:
                                        pass

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
                                review['Review_Link'] = rev_link

                                # review face
                                try:
                                    header = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.left-header")))
                                    div = wait(header, 2).until(EC.presence_of_element_located((By.TAG_NAME, "div")))
                                    attr = div.get_attribute('class')
                                    if 'smiley_smile' in attr:
                                        review['Face_Smiley'] = 1
                                        review['Face_Ok'] = ''              
                                        review['Face_Sad'] = ''    
                                    elif 'smiley_ok' in attr:
                                        review['Face_Smiley'] = ''
                                        review['Face_Ok'] = 1              
                                        review['Face_Sad'] = ''  
                                    elif 'smiley_cry' in attr:
                                        review['Face_Smiley'] = ''
                                        review['Face_Ok'] = ''              
                                        review['Face_Sad'] = 1
                                    else:
                                        review['Face_Smiley'] = ''
                                        review['Face_Ok'] = ''              
                                        review['Face_Sad'] = ''
                                except:
                                    pass
                             
                                # review info
                                review['Dining_Method'] = ''
                                review['Meal_Type'] = ''
                                review['Recommended_Dishes'] = ''
                                review['Visit_Date'] = ''
                                review['Spending_Per_Head'] = ''
                                review['Waiting_Time'] = ''
                                try:
                                    elems = wait(sec, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "section.info.info-row")))
                                    for elem in elems:
                                        divs = wait(elem, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "div")))
                                        if len(divs) == 2 and 'Dining Method' in divs[0].get_attribute('textContent'):
                                            review['Dining_Method'] = divs[1].get_attribute('textContent')
                                        elif len(divs) == 2 and 'Type of Meal' in divs[0].get_attribute('textContent'):
                                            review['Meal_Type'] = divs[1].get_attribute('textContent')   
                                        elif len(divs) == 2 and 'Recommended Dishes' in divs[0].get_attribute('textContent'):
                                            review['Recommended_Dishes'] = divs[1].get_attribute('textContent') 
                                        elif len(divs) == 2 and 'Date of Visit' in divs[0].get_attribute('textContent'):
                                            review['Visit_Date'] = divs[1].get_attribute('textContent')
                                        if 'day(s) ago' in review['Visit_Date']:
                                            try:
                                                num = int(review['Visit_Date'].replace('day(s) ago', '').strip())
                                                review['Visit_Date'] = datetime.today() - timedelta(days=num)
                                                review['Visit_Date'] = date.strftime('%Y-%m-%d')
                                            except:
                                                pass
                                        elif len(divs) == 2 and 'Spending Per Head' in divs[0].get_attribute('textContent'):
                                            review['Spending_Per_Head'] = divs[1].get_attribute('textContent') 
                                        elif len(divs) == 2 and 'Waiting Time' in divs[0].get_attribute('textContent'):
                                            review['Waiting_Time'] = divs[1].get_attribute('textContent')
                                except:
                                    pass

                                # detailed rating
                                try:
                                    div = wait(sec, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "section[class='sr2-review-list2-detailed-rating-section detail']")))
                                    elems = wait(div, 2).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.subject")))
                                    for elem in elems:
                                        name = wait(elem, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div.name"))).get_attribute('textContent').strip()
                                        score = 0
                                        spans = wait(elem, 2).until(EC.presence_of_all_elements_located((By.TAG_NAME, "span")))
                                        for span in spans:
                                            attr = span.get_attribute('class')
                                            if 'yellowstar' in attr:
                                                score += 1

                                        review[name + '_Score'] = score
                                except:
                                    pass
                                review['Extraction Date'] = stamp
                                reviews = reviews.append([review.copy()])
                                nrevs += 1
                                if nrevs == rev_limit:
                                    end = True
                                    break
                            except:
                                pass 
                        # moving to the next page
                        if end: break
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
    if data.shape[0] > 0:
        data['Extraction Date'] = pd.to_datetime(data['Extraction Date'])
    if reviews.shape[0] > 0:
        reviews['Extraction Date'] = pd.to_datetime(reviews['Extraction Date'])
    df1 = pd.read_excel(output1)
    df2 = pd.read_excel(output2)
    df1 = df1.append(data)
    df2 = df2.append(reviews)
    df1.to_excel(output1, index=False)
    df2.to_excel(output2, index=False)
 
def get_restaurants_links(driver, urls, settings):

    res_limit = settings['Restaurants Limit']
    links = {}
    
    for i, url in enumerate(urls):
        driver.get(url)
        exit = False
        links[f"Search Link {i+1}"] = []
        nres = 0
        while True:
            try:
                restaurants = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "h2.title-name")))
                for res in restaurants:
                    try:
                        link = wait(res, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                        if 'restaurants?chainId=' in link: continue
                        links[f"Search Link {i+1}"].append(link)
                        nres += 1
                        if nres == res_limit:
                            exit = True
                            break
                    except:
                        pass

                if exit: break

                # moving to the next results page
                try:
                    url = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='pagination-button next js-next']"))).get_attribute('href')
                    driver.get(url)
                except:
                    break
            except:
                break

    return links



def search_restaurants(driver, res_search, settings):

    res_limit = settings['Restaurants Limit']
    results = []
    for keywords in res_search:
        name = keywords[0]
        loc = keywords[1]
        if name != '' and loc != '':
            driver.get(f'https://www.openrice.com/en/hongkong/restaurants?what={name}&where={loc}')
        elif name != '' and loc == '':
            driver.get(f'https://www.openrice.com/en/hongkong/restaurants?what={name}')
        elif loc != '' and name == '':
            driver.get(f'https://www.openrice.com/en/hongkong/restaurants?where={loc}')

        exit = False
        nres = 0
        while True:
            try:
                restaurants = wait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "h2.title-name")))
                for res in restaurants:
                    try:
                        link = wait(res, 2).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                        if 'restaurants?chainId=' in link: continue
                        results.append((link, name, loc))
                        nres += 1
                        if nres == res_limit:
                            exit = True
                            break
                    except:
                        pass

                if exit: break

                # moving to the next results page
                try:
                    url = wait(driver, 2).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a[class='pagination-button next js-next']"))).get_attribute('href')
                    driver.get(url)
                except:
                    break
            except:
                break

    return results

def initialize_output(stamp):

    
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
    workbook1.add_worksheet()
    workbook1.close()    
    workbook2 = xlsxwriter.Workbook(output2)
    workbook2.add_worksheet()
    workbook2.close()

    return output1, output2

def get_inputs():
 
    print('Processing The Settings Sheet ...')
    print('-'*75)
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\openrice_settings.xlsx'
    else:
        path += '/openrice_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "openrice_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        res_urls, res_search, search_urls = [], [], []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            name, loc = '', ''
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Restaurant Link':
                    if 'restaurants' in row[col]:
                        search_urls.append(row[col])
                    else:
                        res_urls.append(row[col])
                elif col == 'Restaurant Name':
                    name = row[col]
                elif col == 'Restaurant Location':
                    loc = row[col]
                else:
                    settings[col] = row[col]

            if name != '' or loc != '':
                res_search.append((name, loc))
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    # checking the settings dictionary
    keys = ["New Restaurants", "Best Rated Restaurants", "Most Popular Restaurants", "Most Bookmarked Restaurants", "Best Dessert Restaurants", "Scrape Reviews", "Reviews Limit", "Restaurants Limit"]
    for key in keys:
        if key not in settings.keys():
            print(f"Warning: the setting '{key}' is not present in the settings file")
            settings[key] = 0
        try:
            settings[key] = int(float(settings[key]))
        except:
            input(f"Error: Incorrect value for '{key}', values must be numeric only, press an key to exit.")
            sys.exit(1)

    return settings, res_urls, res_search, search_urls

def main():
   
    freeze_support()
    start = time.time()
    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    settings, res_urls, res_search, urls = get_inputs()
    output1, output2 = initialize_output(stamp)
    homepages = ["https://www.openrice.com/en/hongkong/chart/best-rating", "https://www.openrice.com/en/hongkong/chart/most-popular", "https://www.openrice.com/en/hongkong/chart/most-bookmarked", "https://www.openrice.com/en/hongkong/chart/best-dessert", "https://www.openrice.com/en/hongkong/new-restaurants"]

    print('Initializing The Bot ...')
    print('-'*75)
    try:
        driver = initialize_bot()
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()
   
    print('Searching The Site By The User Input Keywords')
    print('-'*75)
    results = search_restaurants(driver, res_search, settings)
    print('Getting The Restaurants Links From The User Input Search URLs')
    if urls:
        search_urls = get_restaurants_links(driver, urls, settings)
        for key, value in search_urls.items():
            homepages.append({key:value})
    if res_urls:
        homepages.append({'User Input':res_urls})
    if results:
        homepages.append({'Search Result':results})

    for page in homepages:
        try:
            scrape_restaurants(driver, output1, output2, page, settings)
        except Exception as err: 
            #print(f'Error:\n')
            #print(str(err))
            driver.quit()
            time.sleep(5)
            driver = initialize_bot()

    driver.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 2)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')

if __name__ == '__main__':

    main()