import os
import pandas as pd
import pytesseract as tess

#date = '202010'
date = input('Input the date in this form 202010 for Oct 2020: ')        
#path = r'Q:\Information Security\Operations\Monthly Reporting\Archer\minh_test\@@2020 OCT'  
path = input('input the folder file')

snow_count = int(input("Enter the snow count : ") or "0")
#snow_count = input("type in the snow count")

malware_people_count = int(input("Enter the the number of malware cleaned by people : ") or "0")
#malware_people_count = input("type in the number of malware cleaned by people")

def find_link(name, path):
    try:
        for root, dirs, files in os.walk(path):
            for filename in files:
                if(name in filename):
                    return path + '\\' + filename
    except:
        return 0
    
#add the part to open excel proofpoint link(blockedlist)
def count_proofpoint(date,proofpoint_link):
    try:
        df_tapcount = pd.read_csv(proofpoint_link)
        count = 0
        for a in df_tapcount['Test']:
            if(date in str(a)):
                count+= 1
        return int(count/2)
    except:
        return 0
#count the scep number
def count_scep(incident_link):
    try:
        scep = pd.read_excel(incident_link)
        scep_count = (scep['Unnamed: 1'][len(scep['Unnamed: 1'])-1])
        return scep_count
    except:
        return 0

#count the backdraft number
def count_backdraft(html_link):
    try:
        #get the sum of backdraft
        backdraft = pd.read_html(html_link)
        backdraft_count = sum(backdraft[0]['Count']) -  backdraft[0]['Count'][0]
        return backdraft_count
    except:
        return 0

#count the Tap total numbr in a picture
def count_tap(picture_link):
    try:
        #picture = 'threat_campaigns_vs_individual_threats-2020_10_01-2020_10_31.png'
        tess.pytesseract.tesseract_cmd = r'Q:\Information Security\Operations\Monthly Reporting\Archer\minh_test\Tesseract-OCR\tesseract.exe'
        from PIL import Image
        img = Image.open(picture_link)
        text = tess.image_to_string(img)
        a = text.split('All Threats')[0].split()
        tap_count = int(a[-1].replace(',',''))
        return tap_count
    except:
        return 0

#small functions to count website
malwarebytes_link = find_link('Malwarebytes', path)
mwb = pd.read_csv(malwarebytes_link)
def web(a):
    if('Website' in a):
        return True
    else:
        return False
def malware(a):
    if('Malware' in a or 'Exploit' in a):
        return True
    else:
        return False

def main():
    website_count = sum(mwb['Category'].apply(lambda x: web(x)))
    malware_count = sum(mwb['Category'].apply(lambda x: malware(x)))
    
    scep_count = count_scep(find_link("Incident", path))
    proofpoint_count = count_proofpoint(date,find_link('blockedlist',path))
    tap_count = count_tap(find_link('png', path)) 
    backdraft_count= count_backdraft(find_link('html', path))
    malware_auto_count= scep_count +malware_count
    event_count = backdraft_count + website_count
    engineering_count = proofpoint_count+ tap_count
    label = ['SCEP Count', 
             "Proofpoint Count", 
             "TAP Count", "Website Count", 
             "Malware + Exploit Count", 
             "Backdraft Count", 
             "SNOW Count", 
             "Malware cleaned by automation", 
             "Number of Events", 
             "Social Engineering Count",
             "Additional Malware cleaned by people"] 
    value =[scep_count,proofpoint_count, tap_count, website_count, malware_count, backdraft_count,
    snow_count, malware_auto_count, event_count, engineering_count, malware_people_count] 

    # Gather information
    archer = pd.DataFrame(data = value, index = label, columns = ['Archer report'])
    print(archer)
    #write to an excel file
    archer.to_excel(path + '\\' + "my_archer_report.xlsx")
    print("It is done! Please check the file my_archer_report.xlsx in the directory you entered. Thank you!")
    
    

if __name__ == "__main__":
    main()

#pip install pytesseract, xlrd, pandas, image, lxml, openpyxl or python -m pip install