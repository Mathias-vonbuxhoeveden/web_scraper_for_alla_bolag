import requests
from bs4 import BeautifulSoup
import numpy as np
import pandas as pd


fin_vars = [ 'Nettoomsättning', 
             'Övrig omsättning', 
             'Rörelseresultat (EBIT)', 
             'Resultat efter finansnetto',
             'Årets resultat',
             'Tecknat ej inbetalt kapital',
             'Anläggningstillgångar',
             'Omsättningstillgångar',
             'Tillgångar',
             'Eget kapital',
             'Obeskattade reserver',
             'Avsättningar (tkr)',
             'Långfristiga skulder',
             'Kortfristiga skulder', 
             'Skulder och eget kapital']




def extract_financial_info(soup_object):
    
    df = soup_object.find_all('td')
    
    values = []
    
    for data in df:
        
        val = data.text.replace('\n','').replace(' ','').strip()
        values.append(val)
        
    values = list(filter(None, values))
        
    return values



def extract_financial_periods(soup_object):
    
    df = soup_object.find_all('th')
    
    values = []
    
    for data in df:
    
        if 'resultaträkning' in data.text.lower():
        
            pass
    
        elif 'nettoomsättning' in data.text.lower():
        
            break
        
        else:
                        
            val = data.text.replace('\n','').replace(' ','').strip()
            values.append(val)
            
    values = list(filter(None, values))
    
    final_list = []
    
    for val in values:
        
        value = val.replace('*','').strip()
        final_list.append(value)
        
    return final_list



def fetch_data(orgnnumber):
    
    url = 'https://www.allabolag.se/{}/bokslut'.format(orgnnumber)
    page = requests.get(url)
    
    soup_object = BeautifulSoup(page.content)
    fin_data = extract_financial_info(soup_object)
    dates = extract_financial_periods(soup_object)
    
    orgs = np.repeat(orgnnumber, len(dates))
    df = pd.DataFrame(data = {'orgnnumber': orgs, 'period': dates})
    
    for i, col in enumerate(fin_vars):
        
        df[col] = fin_data[(i*len(dates)):(i*len(dates) + len(dates))]
        
    return df


def main():

	print("Starting process!")

	df = pd.read_excel('orgnumbers_of_interest.xlsx')
	df.orgnumbers = df.orgnumbers.astype(str)

	res_df = []	

	for org in df.orgnumbers:   

		print("Fetching and parsing data for {}".format(org))
		fin_info = fetch_data(org)    	
		res_df.append(fin_info)    

	final_df = pd.concat(res_df)    
	final_df.to_excel('output_data.xlsx')    
	print("Process finished!")


if __name__ == "__main__":

	main()