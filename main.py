import os
import time
import traceback

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import Tools
import pandas as pd

class Automation():
    def __init__(self):
        self.MP_list = ["NA"]
        self.API_USER, self.API_PWD = "flx-tr1","tr1"
        self.fluxo_creds = (self.API_USER, self.API_PWD)
        user=os.getlogin()
        self.ticket_list=[]
        self.cmd="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=adrese"
        self.cmd1="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=adres"
        self.cmd2="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=iptal"
        self.cmd3="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=iptale"
        self.cmd4="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=alamıyor"
        self.cmd5="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=adresine"
        self.cmd6="https://ticket-api.amazon.com/tickets/?category=Website&type=Customer Service TR Andon Cord&status=Assigned&group=Andon-TR-RBS&assigned_group=Andon-TR-RBS;AndonHV-TR-RBS&text=siparişinizde hata oluştu"




    def selenium(self):
        profile_name = 'Mail_Automation'
        profile_path = rf'{os.path.expanduser("~")}\Chrome Profiles\{profile_name}'
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument(f'user-data-dir={profile_path}')
        chrome_options.add_argument(f'--profile-directory={profile_name}')
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
        chrome_options.add_experimental_option("detach", True)
        browser = webdriver.Chrome( options=chrome_options)
        browser.maximize_window()
        return browser

    def removeFile(self):
        """Remove the files in the downloads directory"""
        file_names=os.listdir("inventoryFolder")
        file_paths=[os.path.abspath(os.path.join("inventoryFolder",file_name)) for file_name in file_names]
        for file in file_paths:
            os.remove(file)


    def ticket_info(self,ticket):
        response = Tools.get_ticket_data(ticket, self.API_USER, self.API_PWD)
        print(response)
        return response['asin']

    def download_fc(self,webdriver,Asin):
        "Checking Alaska for stock"
        #Check Retail FC
        print("Checking if there are retail FC's with stock")
        webdriver.get(f'https://denali-website-eu.aka.amazon.com/inventory-analysis?marketplaceID=712115121&merchantID=54402072512&iaid={Asin}&fnsku={Asin}&merchantSKU={Asin}&view=summary&subview=fc')
        WebDriverWait(webdriver, 100).until(EC.presence_of_element_located((By.XPATH, "//table/tbody/tr/td")))
        time.sleep(5)
        alaska = webdriver.find_element(By.XPATH, "//table/tbody/tr/td")
        status=''
        try:
            assert 'No data to display' in alaska.text
            print('No Available Supply for EU')
        except:
            webdriver.find_element(By.CLASS_NAME,'download-table-button').click()
            time.sleep(2)
            os.rename('inventoryFolder/inventory-summary-table.csv','inventoryFolder/EU FC.csv')
            status+='EU'
            print('FC report downloaded')

        return status

    def check_available_fc(self,retail_fcs,fc_locations):
        #First Check TR FC's second check EU FC's
        self.fc_software=False
        fc_name=None
        max_supply=0

        if 'EU' in fc_locations:
            max_supply = 0
            df = pd.read_csv('inventoryFolder/EU FC.csv')
            for index, row in df.iterrows():
                if float(row['Net Supply']) > float(max_supply) and row['Supply Chain Node Types'] == 'Fulfillment Center' and row['FC']!= 'DTM3':
                    fc_name=row['FC']
                    max_supply=row['Net Supply']
                    self.fc_software = True
                elif float(row['Net Supply']) > float(max_supply) and row['Supply Chain Node Types'] == 'Third Party Logistics' and self.fc_software == False:
                    fc_name=row['FC']
                    max_supply = row['Net Supply']
                    self.fc_software = False

                elif float(row['Net Supply']) > float(max_supply) and row['Supply Chain Node Types'] == 'Dropship':
                    self.is_dropship='STOP'



        if fc_name == 'XTRD' or fc_name == 'XTRC'or fc_name == 'XTRB'or fc_name == 'XTRA':
            self.fc_software = False

        if fc_name=='IST2':
            self.fc_software = True
        print(fc_name,max_supply)
        return fc_name, max_supply

    def create_ticket(self,fc_location,fc_group,asin,fc_id):
        description_retail=f"""Hi, 

This is regarding an Andon Cord. 
Issues: Item_pacakge_dimensions and item_package_weight is incorrect.

Order ID : order_i_d
FCSKU: f_c_s_k_u
ASIN: {asin} 

______________________________________________________
Action_required

Calculate the item package dimensions and weight in the cubicscanner.

Regards,"""
        description_tr=f"""ASIN: {asin}

Merhaba,

Ürünü cubiscan'den geçirip, bilgilerini aşağıdaki formatta "Metre" uzunluk biriminden iletmenizi rica ederim.

Width:
Depth:
Height:
Weight:

Teşekkürler,
Arif         
"""

        if fc_id=='XTRA' or fc_id=='XTRB' or fc_id=='XTRC':
            data = {'category': 'External Fulfillment - TR', 'type': fc_id, 'item': f'{fc_id} - AndonBinCheck',
                    'assigned_group': fc_group, 'asin': asin, 'requester_login': 'garifemr',
                    'short_description': f'Andon Cord Cubicscan Escalation - {asin}', 'details': description_tr,
                    'impact': 3}

            response = Tools.create_ticketV2(data=data, fluxo_user=self.API_USER, fluxo_pwd=self.API_PWD)

        elif fc_id=='XTRD':
            data = {'category': 'External Fulfillment - TR', 'type': fc_id, 'item': 'PO Check',
                    'assigned_group': fc_group, 'city': fc_location, 'asin': asin, 'requester_login': 'garifemr',
                    'building': fc_id,
                    'short_description': f'Andon Cord Cubicscan Escalation - {asin}', 'details': description_tr,
                    'impact': 3}

            response = Tools.create_ticketV2(data=data, fluxo_user=self.API_USER, fluxo_pwd=self.API_PWD)
        else:
            data = {'category': 'iss', 'type': 'Retail Request', 'item': 'Cubiscan Escalation','asin':asin,
                    'assigned_group': fc_group, 'building':fc_id, 'requester_login': 'garifemr',
                    'city':fc_location,
                    'short_description': f'Andon Cord Cubicscan Escalation - {asin}', 'details': description_retail,
                    'impact': 3}

            response = Tools.create_ticketV2(data=data, fluxo_user=self.API_USER, fluxo_pwd=self.API_PWD)
        return response

    def correspond_main(self,webdriver,ticket_link,description):

        #Update main ticket
        webdriver.get(f'https://t.corp.amazon.com/{ticket_link}/communication')

        WebDriverWait(webdriver, 100).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="sim-communicationActions--createComment"]')))

        #fill description
        webdriver.find_element(By.XPATH,'//*[@id="sim-communicationActions--createComment"]').send_keys(description)
        time.sleep(3)

        #click correspondance
        webdriver.find_element(By.XPATH,'//*[text()="Post to Correspondence"]').click()

    def resolve_immediately(self,webdriver):
        WebDriverWait(webdriver, 20).until(
            EC.presence_of_element_located((By.XPATH,'//*[text()="Status"]')))
        webdriver.find_element(By.XPATH,'//*[text()="Status"]').click()
        WebDriverWait(webdriver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[text()="Resolved"]')))
        webdriver.find_element(By.XPATH,'//*[text()="Resolved"]').click()
        WebDriverWait(webdriver, 5).until(
            EC.presence_of_element_located((By.ID, 'closure-code-field')))
        webdriver.find_element(By.ID, 'closure-code-field').click()
        WebDriverWait(webdriver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[text()="Successful"]')))
        webdriver.find_element(By.XPATH, '//*[text()="Successful"]').click()
        webdriver.find_element(By.ID,'root-cause-field').click()
        webdriver.find_element(By.XPATH,'//*[@aria-label="Filter root causes"]').send_keys("Inappropriate Pull-Delivery/Delayed/Pricing/Promotion Error")
        WebDriverWait(webdriver, 5).until(EC.presence_of_element_located((By.XPATH, '//*[@title="Inappropriate Pull-Delivery/Delayed/Pricing/Promotion Error"]')))
        webdriver.find_element(By.XPATH, '//*[@title="Inappropriate Pull-Delivery/Delayed/Pricing/Promotion Error"]').click()
        webdriver.find_element(By.XPATH, '//*[text()="Resolve"]').click()

    def update_ticket_status(self,webdriver):
        WebDriverWait(webdriver, 20).until(
            EC.presence_of_element_located((By.XPATH,'//*[text()="Status"]')))
        webdriver.find_element(By.XPATH,'//*[text()="Status"]').click()
        WebDriverWait(webdriver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[text()="Pending"]')))
        webdriver.find_element(By.XPATH,'//*[text()="Pending"]').click()
        WebDriverWait(webdriver, 5).until(
            EC.presence_of_element_located((By.ID, 'pending-reason-field')))
        webdriver.find_element(By.ID, 'pending-reason-field').click()
        WebDriverWait(webdriver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[text()="Inbound Bin Check"]')))
        webdriver.find_element(By.XPATH, '//*[text()="Inbound Bin Check"]').click()
        webdriver.find_element(By.XPATH, '//*[text()="Save"]').click()

    def get_csi_selenium(self, browser, asin):
        csi_link = f'https://csi.amazon.com/view?view=blame_o&item_id={asin}&marketplace_id=338851&customer_id=&merchant_id=&sku=&fn_sku=&gcid=&fulfillment_channel_code=&listing_type=purchasable&submission_id=&order_id=&external_id=&search_string=&realm=USAmazon&stage=prod&domain_id=&keyword=item_package_&submit=Show'
        browser.get(csi_link)
        WebDriverWait(browser, 60).until(EC.presence_of_element_located((By.CLASS_NAME, "CSIForm")))
        merchant=browser.find_element(By.XPATH,'//*[@id="productdata"]/tbody/tr[1]/td[8]/div/a').text
        return merchant

    def process_automation(self):
        webdriver=self.selenium()
        ticket_set=Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd,msg=None)
        ticket_set.update(Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd1,msg=None))
        ticket_set.update(Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd2,msg=None))
        ticket_set.update(Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd3,msg=None))
        ticket_set.update(Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd4,msg=None))
        ticket_set.update(Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd5,msg=None))
        ticket_set.update(Tools.get_tickets(self.API_USER,self.API_PWD,cmd=self.cmd6,msg=None))



        ticket_list=list(ticket_set)
        #ticket_list.append('V1039883178')
        retail_fcs=Tools.retail_fc_details()
        print('The tickets to be updated are;', len(ticket_list))
        count=0
        resolved_count=0
        for ticket in ticket_list:
            try:
                self.is_dropship = 'continue'
                print(f'\nWorking on ticket {ticket}, updated {count} tickets and resolved {resolved_count} dropper/deliverable ticket.')
                self.removeFile()
                asin=self.ticket_info(ticket)
                fc_locations=self.download_fc(webdriver,asin)
                if len(fc_locations)>1:
                    fc_name,supply=self.check_available_fc(retail_fcs.keys(),fc_locations)
                    if self.is_dropship=='continue':
                        if supply>0:
                            #Checking FC SOFTWARE
                            merchant=self.get_csi_selenium(webdriver,asin)
                            if merchant.lower()=='fc software' and self.fc_software==False:
                                Data = {'assigned_individual': 'garifemr'}
                                Tools.update_ticket(ticket, Data, self.API_USER, self.API_PWD)
                                self.correspond_main(webdriver, ticket,
                                                     "The attributes are reflecting from FC Software. There are no stock in Retail FC's to correct the attribute.")
                                self.resolve_immediately(webdriver)
                                print('FC SOFTWARE RESOLVED')
                                resolved_count += 1
                                continue
                            #Creating Ticket
                            print('retail_fcs[fc_name][0] is',retail_fcs[fc_name][0],'retail_fcs[fc_name][1] is',retail_fcs[fc_name][1],'asin and FC name is ',asin,fc_name)
                            fc_ticket_id=self.create_ticket(retail_fcs[fc_name][0],retail_fcs[fc_name][1],asin,fc_name)
                            if fc_ticket_id != None:
                                #Updating Main Ticket
                                description=f'Cubiscan ticket has been raised to {fc_name}. CB:  https://tt.amazon.com/{fc_ticket_id} \nWinning merchant is {merchant} \nCSI:  https://csi.amazon.com/view?view=blame_o&item_id={asin}&marketplace_id=338851&customer_id=&merchant_id=&sku=&fn_sku=&gcid=&fulfillment_channel_code=&listing_type=purchasable&submission_id=&order_id=&external_id=&search_string=&realm=USAmazon&stage=prod&submit=Show&domain_id=&keyword=item_package_dimensions'
                                #Data={'assigned_individual':'garifemr','tags':'CB_TR'}
                                if self.fc_software:
                                    Data = {'tags': 'CB_EU','correspondence':description,'pending_reason': 'requester information - fc actionable'}
                                else:
                                    Data = {'tags': 'CB_TR','correspondence':description,'pending_reason': 'requester information - fc actionable'}

                                Tools.update_ticket(ticket,Data,self.API_USER,self.API_PWD)
                                #self.correspond_main(webdriver,ticket,description)
                                #self.update_ticket_status(webdriver)
                                print('Ticket Created Succesfully and main ticket has been updated')
                                count+=1
                            else:
                                print('Failed to create Fc ticket moving on')
                        else:
                            print('Not Enough Supply to CB')
                            self.correspond_main(webdriver, ticket, 'There are no stocks left in FC. Asin deliverable')
                            Data = {'assigned_individual': 'garifemr','tags':'CB_NN'}
                            Tools.update_ticket(ticket, Data, self.API_USER, self.API_PWD)
                            self.resolve_immediately(webdriver)
                            print('No FC to open CB, resolved ticket as no issue')
                            resolved_count += 1
                    else:
                        Data = {'assigned_individual': 'garifemr','tags':'CB_NN'}
                        Tools.update_ticket(ticket, Data, self.API_USER, self.API_PWD)
                        self.correspond_main(webdriver, ticket, "There is only stock in DropperShiperNode. DropperShipperNode is out of Andon Cord's scope for delivery issues.")
                        self.resolve_immediately(webdriver)
                        print('Dropshippernode ticket')
                        resolved_count += 1
                else:
                    self.correspond_main(webdriver, ticket, 'No stock in FC. Asin deliverable')
                    Data = {'assigned_individual': 'garifemr','tags':'CB_NN'}
                    Tools.update_ticket(ticket, Data, self.API_USER, self.API_PWD)
                    self.resolve_immediately(webdriver)
                    print('No FC to open CB, resolved ticket as no issue')
                    resolved_count += 1

            except:
                traceback.print_exc()

        print(f'\nUpdated {count} tickets and resolved {resolved_count} ticket.')
        webdriver.quit()

if __name__ == "__main__":
    obj = Automation()
    obj.process_automation()
