
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


st_erdc= [*CONFIDENTIAL*]
df_erdc=["","","","","","","","","","","","","","",""]

st_b1= [*CONFIDENTIAL*]
df_b1=["","","","","","","","","","","","","","","",""]

counter_a=0
counter_b=0
counter_c=0
counter_d=0

prog_erdc= '#'
prog_b1= '#'
stat = ['Loading...','Finished']

browser = webdriver.Chrome(*PATH-CONFIDENTIAL*)
browser.get(('https://www.dell.com/support/home/uk/en/ukdhs1/'))

print('################## Downloading ##################')

for x in st_erdc:
    #print(st[counter])

    username = WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.NAME, "entry-main-input")))
    username.send_keys(st_erdc[counter_a])
    
    
    searchButton = browser.find_element_by_xpath('//*[@id="home-product-box"]/entry-selection/div[1]/div/form/div[1]/div[1]/div/span/button')
    searchButton.click()
    
    sysconfig = WebDriverWait(browser, 4).until(
    EC.presence_of_element_located((By.ID, "systemconfiguration")))
    sysconfig.click()
    
    export = WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.ID, "exporticon")))
    export.click()
    
    browser.get(('https://www.dell.com/support/home/uk/en/ukdhs1/'))
    
    #print (counter_a)
    counter_a +=1
    prog_erdc += '#'
    per_a=(counter_a/15)*100
    print('ERDC_Config Downloading ',str(stat[0]) + '[',prog_erdc,'] ',str(round(per_a,0)) + '%', end='\r')
    
    if counter_a == 15: 
        print('ERDC_Config Downloading ',str(stat[1]) + '[',prog_erdc,'] ',str(round(per_a,0)) + '%', end='')
        break
        
print('\n ERDC Download Complete')

for x in st_b1:
    #print(st[counter])
    
    
    username = WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.NAME, "entry-main-input")))
    username.send_keys(st_b1[counter_b])
    
    
    searchButton = browser.find_element_by_xpath('//*[@id="home-product-box"]/entry-selection/div[1]/div/form/div[1]/div[1]/div/span/button')
    searchButton.click()
    
    sysconfig = WebDriverWait(browser, 4).until(
    EC.presence_of_element_located((By.ID, "systemconfiguration")))
    sysconfig.click()
    
    export = WebDriverWait(browser, 3).until(
    EC.presence_of_element_located((By.ID, "exporticon")))
    export.click()
    
    browser.get(('https://www.dell.com/support/home/uk/en/ukdhs1/'))
    
    #print (counter_b)
    counter_b +=1
    prog_b1 += '#'
    per_b=(counter_b/16)*100
    print('B1_Config Downloading ',str(stat[0]) + '[',prog_b1,'] ',str(round(per_b,0)) + '%', end='\r')
    
    if counter_b == 16:
        print('B1_Config Downloading ',str(stat[1]) + '[',prog_b1,'] ',str(round(per_b,0)) + '%', end='')
        break
        
print('\n Building 1 Download Complete')
    
        
print('################## WRITING ##################')
    
with ExcelWriter('test.xlsx') as writer:
    for x in st_erdc:
        
        df_erdc[counter_c]=pd.read_csv(*PATH-CONFIDENTIAL*+str(st_erdc[counter_c]) +'.csv')
        
        df_erdc[counter_c].to_excel(writer, sheet_name=str(st_erdc[counter_c]))
        
        #print(counter_c)
        counter_c +=1
        if counter_c==14:
                break
                
    print('ERDC Writing Complete')

    for x in st_b1:

        df_b1[counter_d]=pd.read_csv(*PATH-CONFIDENTIAL*+str(st_b1[counter_d]) +'.csv')

        df_b1[counter_d].to_excel(writer, sheet_name=str(st_b1[counter_d]))

        #print(counter_d)
        counter_d +=1
        if counter_d==15:
            break
            
writer.save()
print('Building 1 Writing Complete')
print('Done! ABOSTO!')
