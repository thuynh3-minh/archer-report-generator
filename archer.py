import os
import pandas as pd
import pytesseract as tess
from tkinter import filedialog
from tkinter import *
import tkinter as tk

#function to return the link for each file
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
        #location of the pytesseract.exe in the computer. Here you have to adjust it depending on where it is stored in your pc
        tess.pytesseract.tesseract_cmd = r'Q:\Information Security\Operations\Monthly Reporting\Archer\minh_test\Tesseract-OCR\tesseract.exe'
        from PIL import Image
        img = Image.open(picture_link)
        text = tess.image_to_string(img)
        a = text.split('All Threats')[0].split()
        tap_count = int(a[-1].replace(',',''))
        return tap_count
    except:
        return 0
    
#small functions to support the website_count and malware_count line 9-10 in the oldmain aka report function. Count numbers in the excel file
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

#main function takes three user inputs from the UI. It takes date, path and number of snow emails and malware emails removed by people in that order
def report(input1,input2,input3,folder_path):

    date = input1
    path = folder_path
    snow_count = input2
    malware_people_count = input3

    #small functions to count website
    malwarebytes_link = find_link('Malwarebytes', path)
    mwb = pd.read_csv(malwarebytes_link)
    
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
    #print(archer)
    
    #write to an excel file in the old directory
    archer.to_excel(path + '\\' + "my_archer_report.xlsx")
    
    #can be improve: create a pop up window to inform success.
    print("It is done! Please check the file my_archer_report.xlsx in the directory you entered. Thank you!")
    

#basic interface of the app
root = tk.Tk()
heightApp=500
widthApp=500
root.geometry(str(heightApp)+"x"+str(widthApp))
root.title("Archer report creator")
root.iconbitmap(r'Q:\Information Security\Operations\Monthly Reporting\Archer\minh_test\archer\archericon.ico')
canvas=tk.Canvas(root, height=heightApp, width=widthApp)
canvas.pack()

#function support the browse button
def browse_button():
    # Allow user to select a directory and store it in global var
    # called folder_path
    global folder_path
    filename = filedialog.askdirectory()
    folder_path.set(filename)
    #print(filename)
    
#a test function to test every button and entry. Activate when button create report is clicked. Now it is replaced by the report function 
def test_function(entry1,entry2,entry3,folder_path):
    entry2=int(entry2 or 0)
    entry3=int(entry3 or 0)
    print(type(entry1))
    print(type(entry2))
    print(type(entry3))
    print(type(folder_path))

frame=tk.Frame(root,bg="#d3d3d3",bd=5)#bg=#
frame.place(relx=0.5,rely=0,relwidth=0.75,relheight=0.6,anchor='n')

#date entry 
label1 = tk.Label(frame,font=30,text="Enter the date in form YYYYMM:",bg="#d3d3d3").place(relx=0,rely=0.10,relwidth=0.65,relheight=0.1)
entry1 = tk.Entry(frame,font=30)
entry1.place(relx=0,rely=0.20,relwidth=0.65,relheight=0.1)

#snowcount entry
label2 = tk.Label(frame,font=30,text="Enter the snow count",bg="#d3d3d3").place(relx=0,rely=0.30,relwidth=0.65,relheight=0.1)
entry2 = tk.Entry(frame,font=30)
entry2.place(relx=0,rely=0.40,relwidth=0.65,relheight=0.1)

#malwarebytes cleaned by people entry
label3 = tk.Label(frame,font=30,text="Enter the Malwarebytes",bg="#d3d3d3").place(relx=0,rely=0.50,relwidth=0.65,relheight=0.1)
entry3 = tk.Entry(frame,font=30)
entry3.place(relx=0,rely=0.60,relwidth=0.65,relheight=0.1)

#browse button and path
folder_path = StringVar()
entry4 = tk.Label(frame,textvariable=folder_path).place(relx=0, rely=0.9,relwidth=0.65,relheight=0.1)
label_file_explorer = tk.Label(frame,text = "Folder location for archer",font=30,bg="#d3d3d3").place(relx=0, rely=0.7,relwidth=0.65,relheight=0.1) 
button2 = tk.Button(frame,text="Browse", command=browse_button,bg="#d3d3d3").place(relx=0, rely=0.8,relwidth=0.65,relheight=0.1)

#create report button
button = tk.Button(frame,text="Create report", font=30, command=lambda: report(entry1.get(),entry2.get(),entry3.get(),folder_path.get()))
button.place(relx=0.7,rely=0.2,relwidth=0.3,relheight=0.1)

root.mainloop()
