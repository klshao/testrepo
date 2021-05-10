while True:
    
    try:
        address = input("Please input the complete address of the file:(for example: F:/Calculations/AllSiMe3/Summary/output/)")
        folder_name = input("Please input the name of the folder: (for example :1-5.out)")
        
        with open(address+folder_name,'r') as output_file:
            pass
        
    except:
        print("Invalid address or folder name, please try again.")
        
    else:
        print("The file was successfully located.")
        break
                    
import os.path
import xlwt
import re   #extract number from a string

outWorkbook = xlwt.Workbook()
outSheet = outWorkbook.add_sheet('Sheet1', cell_overwrite_ok=True)

outSheet.write(0,0,"Chemicals")
outSheet.write(0,1,"UCCSD(T)")
outSheet.write(0,2,"ZPE")
outSheet.write(0,3,"Electronic energy + ZPE(0 K)")
outSheet.write(0,4,"Electronic energy + Thermal energy correction")
outSheet.write(0,5,"Electronic energy + Thermal enthalpy correction")
outSheet.write(0,6,"Electronic energy + Gibbs Free Energy correction")
                    
with open(address+folder_name,'r') as output_file:
    
#For address here, could "/" only be used for Windows system, '\' will not be understood. Just a primary version, 
#further polish is needed to filter "\"!
    
    lines = output_file.readlines()
    index = 0
    folder_number = 1
    folder_number_string = '1'
    file_name = ''
    file_created = False
    length_of_the_file = len(lines)
    number_extracted = ''
    chemical_name_list = []
    chemical_name = ''                   #naming the chemical in the .xls file
    
    while index < length_of_the_file -1:
        
        while True:                                          #"Normal termination of Gaussian 09" not in lines[index]
            
            #if "%chk=" in lines[index]:
                
            if "Initial command:" in lines[index]:
                    file = open(os.path.join(address, 'temp_file.out'),'w')
                    file_created = True
            
            if "%chk=" in lines[index]:
                
                file_name = folder_number_string + '_' + lines[index][6:-5] + '.out'
                complete_name = os.path.join(address, file_name)
                chemical_name = folder_number_string + '_' + lines[index][6:-5]
                folder_number = folder_number + 1
                folder_number_string = str(folder_number)
                print(f'File: {file_name} has been created! at {address}')
                chemical_name_list = [chemical_name]
                outSheet.write(folder_number-1,0,chemical_name_list[0])
                
                
            if file_created:
                file.write(lines[index])

            if "Zero-point correction=" in lines[index]:             #ccsd(t)
                number_extracted = re.findall(r"\d+\.?\d*",lines[index])    #extract number from a string
                outSheet.write(folder_number-1,2,float(number_extracted[0]))   #write data into .xls file 
                
            if "CCSD(T)= " in lines[index]:             #ccsd(t)
                chemical_name_list = [chemical_name]
                number_extracted = re.findall(r"\d+\.?\d*",lines[index])    #extract number from a string
                ccsd_num = float('-'+number_extracted[0][:-2])*1000
                outSheet.write(folder_number-1,1,ccsd_num)   #write data into .xls file 
                
            if "Sum of electronic and zero-point Energies=" in lines[index]:
                chemical_name_list = [chemical_name]
                number_extracted = re.findall(r"\d+\.?\d*",lines[index])    #extract number from a string
                outSheet.write(folder_number-1,3,float('-'+number_extracted[0]))   #write data into .xls file 
            
            if "Sum of electronic and thermal Energies=" in lines[index]:
                chemical_name_list = [chemical_name]
                number_extracted = re.findall(r"\d+\.?\d*",lines[index])    #extract number from a string
                outSheet.write(folder_number-1,4,float('-'+number_extracted[0]))
                
            if "Sum of electronic and thermal Enthalpies=" in lines[index]:
                chemical_name_list = [chemical_name]
                number_extracted = re.findall(r"\d+\.?\d*",lines[index])    #extract number from a string
                outSheet.write(folder_number-1,5,float('-'+number_extracted[0]))
                
            if "Sum of electronic and thermal Free Energies=" in lines[index]:
                chemical_name_list = [chemical_name]
                number_extracted = re.findall(r"\d+\.?\d*",lines[index])    #extract number from a string
                outSheet.write(folder_number-1,6,float('-'+number_extracted[0]))
        
            index = index + 1
            
            if index == length_of_the_file -1:
                break
            
            elif ("Normal termination of Gaussian 09" in lines[index]) and ("Initial command:" in lines[index+1]):
                break
                
        file.close()
        os.replace(os.path.join(address, 'temp_file.out'), complete_name)
        
        with open(complete_name,'a') as last_word:
            last_word.write(lines[index])                               #File created
            
        file_created = False
        index = index + 1

outWorkbook.save(address + folder_name[:-4] +'_Output_summary.xls')
print("The summary Excel file has also been successfully created!") 