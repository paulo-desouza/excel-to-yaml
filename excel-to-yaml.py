from excel_methods import *
import yaml


### We need to take the data from the form, and save it in a YAML file. 
 
### Then, with the YAML files, we can easily retrieve the information
### to generate different types of reporting (weekly, monthly)



### TO-DO ###

    # CHECK FOR THE VALIDITY OF THE WEEK VALUE
    # CHECK FOR THE VALIDITY OF THE OTHER DATA AS WELL
    # CONVERT THE WEEK VALUE
    
    # DONT TAKE THE TOTAL FROM THE SPREADSHEET, CALCULATE IT YOURSELF
    

file_list = get_file_names()

for file in file_list:
    
    if ".xlsx" in file:
        
        try:
        
            wb = load_workbook(file)
            ws = wb.active
            
            
            
            # Run Checks to see if there is data in the forms!
            cannot_be_empty = ["F2", "I2",
                               "B4", "D4", "F4", "H4", "J4",
                               "B5", "D5", "B6", "D6", 
                               "B9", "D9", "E9", "F9", "G9", "H9", "I9",
                               "B10", "D10", "E10", "F10", "G10", "H10", "I10",
                               "B11", "D11", "E11", "F11", "G11", "H11", "I11", 
                               "B12", "D12", "E12", "F12", "G12", "H12", "I12", 
                               "B15", "D15", "F15", "H15", "J15",
                               "B16", "D16", "B17", "D17",
                               "B20", "D20", "E20", "F20", "G20", "H20", "I20",
                               "B21", "D21", "E21", "F21", "G21", "H21", "I21",
                               "B22", "D22", "E22", "F22", "G22", "H22", "I22",
                               "B23", "D23", "E23", "F23", "G23", "H23", "I23",
                               "B25", "D25", "F25",
                               "B27", "C27", "D27", "E27", "F27", "G27",
                               "B28"
                               ]
            
            needs_to_be_string = []
            needs_to_be_int = []
            
            
            isnt_int =      []
            isnt_string =   []
            empty =         []
            
            for cell in cannot_be_empty:
                if ws[cell].value == None:
                    empty.append(cell)
                    
            
            
            if len(empty) == 0:
                data = {}
                
                data["EA"] = ws["F2"].value
                data["Week"] = ws["I2"].value
                
                data["Table1"] = {
                    "Celebree": ws["B4"].value,
                    "District Goal": ws["D4"].value,
                    "FTEs Goal": ws["F4"].value,
                    "District Actual": ws["H4"].value,
                    "FTEs Actual": ws["J4"].value,
                    
                    "Weekly Goal": ws["B5"].value,
                    "FTEs Starting Date (Goal)": ws["D5"].value,
                    
                    "Actual FTEs": ws["B6"].value,
                    "FTEs Starting Date (Actual)": ws["D6"].value,
                    
                    "Inbound Call Goal & Actuals" : [
                         ws["B9"].value, ws["D9"].value, ws["E9"].value, ws["F9"].value, ws["G9"].value, ws["H9"].value,  ws["I9"].value,  ws["J9"].value ],
                    "Outbound Call Goal & Actuals" : [ws["B10"].value, ws["D10"].value, ws["E10"].value, ws["F10"].value, ws["G10"].value,
                                                    ws["H10"].value,  ws["I10"].value,  ws["J10"].value ],
                    "Visit Schedule Goal & Actuals" : [ws["B11"].value, ws["D11"].value, ws["E11"].value, ws["F11"].value, ws["G11"].value,
                                                    ws["H11"].value,  ws["I11"].value,  ws["J11"].value ],
                    "Enrolled Goal & Actuals" : [ws["B12"].value, ws["D12"].value, ws["E12"].value, ws["F12"].value, ws["G12"].value,
                                                    ws["H12"].value,  ws["I12"].value,  ws["J12"].value ]
                    }
                
                
                
                data["Table2"] = {
                    "Caliday": ws["B15"].value,
                    "District Goal": ws["D15"].value,
                    "FTEs Goal": ws["F15"].value,
                    "District Actual": ws["H15"].value,
                    "FTEs Actual": ws["J15"].value,
                    
                    "Weekly Goal": ws["B16"].value,
                    "FTEs Starting Date (Goal)": ws["D16"].value,
                    
                    "Actual FTEs": ws["B17"].value,
                    "FTEs Starting Date (Actual)": ws["D17"].value,
                    
                    
                    "Inbound Call Goal & Actuals" : [
                         ws["B20"].value, ws["D20"].value, ws["E20"].value,
                         ws["F20"].value, ws["G20"].value, ws["H20"].value,
                         ws["I20"].value,  ws["J20"].value],
                    "Outbound Call Goal & Actuals" : [
                         ws["B21"].value, ws["D21"].value, ws["E21"].value, 
                         ws["F21"].value, ws["G21"].value, ws["H21"].value,  
                         ws["I21"].value,  ws["J21"].value 
                         ],
                    "Visit Schedule Goal & Actuals" : [
                         ws["B22"].value, ws["D22"].value, ws["E22"].value, 
                         ws["F22"].value, ws["G22"].value, ws["H22"].value,
                         ws["I22"].value,  ws["J22"].value
                         ],
                    "Enrolled Goal & Actuals" : [
                         ws["B23"].value, ws["D23"].value, ws["E23"].value, 
                         ws["F23"].value, ws["G23"].value, ws["H23"].value,
                         ws["I23"].value,  ws["J23"].value 
                         ]
                    }
                
                
                
                data["Table3"] = {
                    "Weekly Challenge": ws["B25"].value,
                    "Goal": ws["D25"].value,
                    "Actual": [ws["F25"].value, ws["G25"].value],
                    "Over Due Tasks": [ws["B27"].value, ws["C27"].value, ws["D27"].value, ws["E27"].value, ws["F27"].value,
                                                    ws["G27"].value,  ws["H27"].value],
                    "Opportunity For Success": ws["b28"].value
                    }
                
                
                
                # Output to YAML
                
                with open("{data[EA]}{data[Week]}.yaml", "w") as file:
                    yaml.dump(data, file, default_flow_style=False)
                
                
            else:
                
                print(f"Necessary fields are still empty! Please fill these out before running the program: {empty}")
        
        
        except InvalidFileException:
            print("Please check that none of the excel spreadsheets of this folder are Open. ")
            
            
            
            
            
            

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    