                                                
import socket #Required for IP address and computer name
import platform #Required for getting os type on Linux & Windows
import time #Required for system time
import os   #Required for deletion of temp files.
import csv  #Required in the Append function.

#Required for the pip install for Linux / Windows
import sys
import subprocess                  
from subprocess import PIPE, run 

#Declare Variables and set them as string "Not Found" by default.
computerName = 'Not Found'
ipAddress = 'Not Found'
macAddress = 'Not Found'
processorModel = 'Not Found'
operatingSystem = 'Not Found'
systemTime = 'Not Found'
internetSpeed = 'Not Found'
activePorts = 'Not Found'

#If Linux machine we run the function for pip install of third party packages for Linux
#Check modules required and install if needed using pip.
def pipLinux():
    # Install pandas
    print("Checking missing modules, please wait.")
    pipPandasCmd = [sys.executable, '-m', 'pip', 'install', 'pandas']  # Using sys.executable for pip install
    result = subprocess.run(pipPandasCmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    if result.returncode == 0:
        print("20% Complete")
         
    else:
        print(f"Failed to install pandas: {result.stderr}")
        print("See help menu for trouble shooting.") #If automatic install of module fails we print instructions for manual installation of module.

    # Install psutil
    pipPsutilCmd = [sys.executable, '-m', 'pip', 'install', 'psutil']  # Using sys.executable for pip install. python -m pip ensures that you are installing the package for the specific Python interpreter you are using.
    result = subprocess.run(pipPsutilCmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    if result.returncode == 0:
        print("40% Complete")
         
    else:
        print(f"Failed to install psutil: {result.stderr}")
        print("See help menu for trouble shooting.") #If automatic install of module fails we print instruction for manual installation of module.

    # Install speedtest
    pipSpeedtestCmd = [sys.executable, '-m', 'pip', 'install', 'speedtest-cli']  # Using sys.executable for pip install. python -m pip ensures that you are installing the package for the specific Python interpreter you are using.
    result = subprocess.run(pipSpeedtestCmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    if result.returncode == 0:
        print("60% Complete")
         
    else:
        print(f"Failed to install speedtest-cli: {result.stderr}")
        print("See help menu for trouble shooting.") #If automatic install of module fails we print instruction for manual installation of module.

    # Install getmac
    pipGetmacCmd = [sys.executable, '-m', 'pip', 'install', 'getmac']  # Using sys.executable for pip install. python -m pip ensures that you are installing the package for the specific Python interpreter you are using.
    result = subprocess.run(pipGetmacCmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    if result.returncode == 0:
        print("80% Complete")
         
    else:
        print(f"Failed to install getmac: {result.stderr}")
        print("See help menu for trouble shooting.") #If automatic install of module fails we print instruction for manual installation of module.

    # Install cpuinfo
    pipCpuinfoCmd = [sys.executable, '-m', 'pip', 'install', 'py-cpuinfo']  # Using sys.executable for pip install
    result = subprocess.run(pipCpuinfoCmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    if result.returncode == 0:
        print("100% Complete. All Modules installed.\n")
         
    else:
        print(f"Failed to install cpuinfo: {result.stderr}")
        print("See help menu for trouble shooting.") #If automatic install of module fails we print instructions for manual installation of module.


    getData()


#If Windows machine we run the pip install of third party packages function for Windows
#Check modules required and install if needed using pip.
#This function runs commands in the command prompt using sub process module, pipe, & run.
def pipWin():
    print("Checking missing modules, please wait." )
    #pipPandas
    pipPandasCmd = 'pip install --upgrade pip'   # Installs pip 
    result = run(pipPandasCmd, stdout=PIPE, stderr=PIPE, universal_newlines=True)
    print (result.stderr)
    if result.returncode == 0:
        print("10% Complete")
    #print (result.stdout, result.stderr)

    #pipPandas
    pipPandasCmd = 'pip install pandas'   # Installs Pandas library 
    result = run(pipPandasCmd, stdout=PIPE, stderr=PIPE, universal_newlines=True)
    print (result.stderr)
    if result.returncode == 0:
        print("20% Complete")
    #print (result.stdout, result.stderr)

    #pipSpeedtest
    pipSpeedtestCmd = 'pip install speedtest-cli'   # Installs speedtest library 
    result = run(pipSpeedtestCmd, stdout=PIPE, stderr=PIPE, universal_newlines=True)
    print (result.stderr)
    if result.returncode == 0:
        print("40% Complete")
    #print (result.stdout, result.stderr)
    

    #pippsutil
    pipPsutilCmd = 'pip install psutil'   # Installs psutil library 
    result = run(pipPsutilCmd, stdout=PIPE, stderr=PIPE, universal_newlines=True)
    print (result.stderr)
    if result.returncode == 0:
        print("60% Complete")
    #print (result.stdout, result.stderr)

    #pipGetmac
    pipGetmacCmd = 'pip install getmac'   # Installs getmac library 
    result = run(pipGetmacCmd, stdout=PIPE, stderr=PIPE, universal_newlines=True)
    print (result.stderr)
    if result.returncode == 0:
        print("80% Complete")
    #print (result.stdout, result.stderr)

    #cpuinfo
    pipCpuinfoCmd = 'pip install py-cpuinfo'   # Installs cpuinfo library 
    result = run(pipCpuinfoCmd, stdout=PIPE, stderr=PIPE, universal_newlines=True)
    print (result.stderr)
    if result.returncode == 0:
        print("100% Complete. All Modules installed.\n")
    #print (result.stdout, result.stderr)

    getData()

#Once getData function is complete the data is stored in variables. eg computerName, ipAddress etc. If it fails the variable will stay as 'Not Found'.
#As the script is running we will print results.
def getData():
    import psutil #Place inside function incase it wasnt previously installed.
    global computerName  # Get Computer Name using socket / variables declared global as we are inside a function.
    computerName = socket.gethostname() #Set variable
    print(computerName)
    global ipAddress
    try:
        # Connect to a public server to get the local IP address
        #Putting this in the 'with' statement ensures the socket is properly close incase of no internet and exception error.
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))  # Ping Google on port 80
            ipAddress = s.getsockname()[0]  # Set the IP address variable
    except Exception as e:
        ipAddress = "Error: " + str(e)

    # Print the local IP address
    print(ipAddress)

    try:
        global macAddress  # Get MAC Address using getmac module
        macAddress = None
        import getmac
        macAddress = (getmac.get_mac_address()) 
        print(macAddress)
    except Exception as e:
        macAddress = "No Internet Connection"
        print(macAddress)

    global processorModel  # Declare processorModel as a global variable
    import cpuinfo

    # Retrieve the CPU information dictionary
    cpu_info = cpuinfo.get_cpu_info()

    # Try to get the processor model from available keys
    processorModel = cpu_info.get('brand_raw') or cpu_info.get('arch_string_raw') or cpu_info.get('model') or "Not Found"

    # Print the processor model
    print(processorModel)


    global operatingSystem  # Get Operating System using platform
    operatingSystem = platform.system() + " " + platform.release() #Set variable operatingSystem
    print(operatingSystem)

    global systemTime  
    t = time.time()    # Get the current time
    print(time.ctime(t))
    systemTime = (time.ctime(t))  #Set variable systemTime
  
    global activePorts  # Get Active Ports using psutil
    activePorts = [] #Set variable activePorts
    for conn in psutil.net_connections():
        if conn.status == psutil.CONN_LISTEN: #Check if the Connection is Listening / We can add ESTABLISHED here if required.
            activePorts.append(conn.laddr.port) #Add the Listening Port to the List.
    print(activePorts)    
    downloadSpeed()
#Seperate function for the try speedtest-cli module. It may fail use try except and finally so we don't break entire script.
def downloadSpeed():
    global internetSpeed  
    print("Attempting to test internet speed, please wait..") #This takes some time so we print that it has begun. 
    try:
        import speedtest
        st = speedtest.Speedtest()         
        download_speed = st.download()
        internetSpeed = download_speed / (1024 * 1024)
        print(internetSpeed)
        
    except Exception as e:
            # If any error occurs in the speedtest block, we print an error message and move on. This can fail due to server side issues or no internet connection.
            # If machine is not connected to the internet variable remains as 'Not Found'.
        print("An error occurred during the speed test: " + str(e) + ". This could be due to server-side network issues, try again later.")
        
    finally: #finally will always execute even if speedtest has errors.
        storeNewData()
#Write the variables to a temporary csv file named newData.csv.
#We import pandas inside our function as it may not already be installed when running the script.
def storeNewData():
    import pandas as pd  # Required for reading writing to csv files
    #print("Cleaned Value of computerName for debugging: " + repr(computerName)) This is print debug the repr shows any hidden chars like \n etc
    # Dictionary to represent a new row of data
    newData = {                        
        "computerName": [computerName],
        "ipAddress": [ipAddress],
        "macAddress": [macAddress],
        "processorModel": [processorModel],
        "operatingSystem": [operatingSystem],
        "systemTime": [systemTime],
        "internetSpeed": [internetSpeed],
        "activePorts": [activePorts]
    }
    new_df = pd.DataFrame(newData)
    new_df.to_csv('newData.csv', mode='w', header=0, index=False)
    #Header =0 means the first row is the header. index=False)  # Prevents pandas from using the first column as the index
    checkForFile()
    #print('Variables are now stored in a csv file named newData.csv') This is a debug print statement
    
#Check for existing csv file or create new one. This function must be placed here. Placing it before storeNewData function will cause failure for unkown reasons in Linux.
def checkForFile():
    def writeHeader():
        header = ['Computer_Name', 'IP_Address', 'Mac_Address', 'Processor_Model', 'Operating_System', 'System_Time', 'Internet_Connection_Speed', 'Active_Ports']
        with open(fileName, mode='w', newline='') as file: #New line fixes the issue of it skipping the 2nd row.
            writer = csv.writer(file)
            writer.writerow(header)
        #print("Headers written. Start append function") #Debug print.
        csvAppend() #Skip search and straight to append function
    fileName = "Midtown_IT.csv"

    fileExist = os.path.isfile(fileName)
    print(fileExist)
    if fileExist == False:
        writeHeader()
    else:
        #print('Skip this and starting get data function') #Debug print 
        searchCsv()

    #Check if previous record exists regarding this machine, check for matching machine name in Midtown_IT.csv file in column 1
def searchCsv():
    import pandas as pd
    print('Searching csv for existing match')             #Use pandas to search the csv for a corresponding row containing the hostname, if no match we write to next empty row, if match
    global rowIndex                                       #we replace row with new data from stored data in variables.
    # Load the dataset
    data = pd.read_csv('Midtown_IT.csv')
    # Find the index of the row(s) where 'computerName's' match.
    matchingRows = data.index[data['Computer_Name'] == computerName].tolist()
    # Check if any matches were found.
    if matchingRows:
        for rowIndex in matchingRows:
            # Adding 1 to convert zero-based index to one-based
            #print(f"Match found at row {rowIndex + 1}:")   #This is a debug print statement
            #print("Starting compare function..")           #This is a debug print statement
            #print(data.iloc[rowIndex])                     #This is a debug print statement  
            matchingRowNum() #If match found rowIndex variable carries the matching row number into the matchingRownNum() and saveCsv() function.
    else:
        print("Computer name '" + computerName + "' not found in the dataset.")
        #print("Starting append csv function")    #This is a debug print statement                                          
        csvAppend()     #If not found we can use append when writing to csv 
# if there is no previous record of this machine append it's data to the Midtown_IT.csv file. and tell the user: "process complete new record added".
def csvAppend():
    import pandas as pd
    savechanges = input("Enter [1] to add " + (computerName) + " to the CSV, or [2] to exit without saving: ") 
    if savechanges == "1":
        new_data = {
            'Computer_Name': [computerName,],
            'IP_Address': [ipAddress,],
            'Mac_Address': [macAddress,],
            'Processor_Model': [processorModel,],
            'Operating_System': [operatingSystem,],
            'System_Time': [systemTime,],
            'nternet_Connection_Speed': [internetSpeed,],
            'Active_Ports': [activePorts,]
        }
        new_df = pd.DataFrame(new_data)
        new_df.to_csv('Midtown_IT.csv', mode='a', header=0, index=False) #Header =0 means the first row is the header. index=False)  # Prevents pandas from using the first column as the index
        print('Append function complete')
        print('CSV file now updated.')
        os.remove('newData.csv') 
    else:
        print("No changes made.")
        os.remove('newData.csv')  #Deletes the newData.csv file
# If there is a matching record, an entry with the same computer name, get row number of this match. Call it (matchingRowNum) function
#Using that row number extract only that row to a temporary csv file named "oldData.csv". This is the old data for comparison to new data
def matchingRowNum():
    import pandas as pd
    print(f"Row index variable checking {rowIndex + 1}:")
    row_index = rowIndex
    # Read the CSV file into a DataFrame
    data = pd.read_csv('Midtown_IT.csv')
    # Extract the specific row as a Series
    single_row_series = data.iloc[row_index]
    # Convert the Series back to a DataFrame (with one row)
    single_row = pd.DataFrame([single_row_series])
    # Save the extracted row to a new CSV file
    single_row.to_csv('oldData.csv', index=False, header=False)
    # Compare the two temporary csv files, oldData.csv and newData.csv .
    with open('oldData.csv', 'r') as oldData, open('newData.csv', 'r') as newData:
        if oldData.read() == newData.read():
            oldData.close()       #Can't delete the csv's if we don't close them here.
            newData.close()
            print("Files are identical no changes made, ending program") #inform user no changes detected.
        else:
            print("Changes detected for " + str(computerName) + " Would you like to save changes to the Midtown_IT CSV file?")  #Give choice to save or exit
            oldData.close()
            newData.close()               
            saveChanges = input('Enter [1] to save changes to the Midtown_IT csv or [2] to exit without saving:')
            if saveChanges =='1': 
                saveCsv()
            elif saveChanges =='2':
                deleteTempFiles()    
            else:
                print ('Invalid response')
                matchingRowNum()
# If yes, re-write the Midtown_IT.csv and replace the old row with the new data stored in the variables
def saveCsv():
    import pandas as pd
    #write the new row to the midtown csv
    print("ComputerName Variable within the saveCsv function debug: " + str(computerName))
    # Define the row number to modify
    row_to_modify = 2 + rowIndex  # "+2" avoids changing the header + its 0 based.
    new_row = [computerName, ipAddress, macAddress,
                processorModel, operatingSystem, systemTime, internetSpeed, activePorts]  # New variable values for the row.
    # Read the CSV file into memory
    with open('Midtown_IT.csv') as inf:
        reader = list(csv.reader(inf.readlines()))
    # Modify the "row to modify variable" row (lists are zero-indexed, hence adding 2 to the variable)
    reader[row_to_modify - 1] = new_row 
    # Write the modified data back to the CSV file
    with open('Midtown_IT.csv', 'w', newline='') as outf:
        writer = csv.writer(outf)
        writer.writerows(reader)
        
        #print("Completed writing updates to CSV file using the saveCsv function") #This is a debug print statement 
        deleteTempFiles()
# Delete both temporary CSV files.
def deleteTempFiles():
    os.remove('newData.csv') 
    os.remove('oldData.csv')
    print("Cleaning up temporary files, finished, program complete.")



def helpWin():
    
    print("User instructions for Midtown.py script\n")
    print("The Midtown.py script is designed to collect information from each computer by running the script from a portable USB thumb drive.")
    print("Follow the instructions below for running the script on Windows.")
    print("The instructions will assume you have loaded both the Midtown_IT.csv and Midtown.py script onto your thumb drive.")
    print("Note this script can be run from within a Python IDE or Windows command prompt.")
    print("We will assume not every PC has a Python IDE installed, so we will only be using the command prompt.\n")

    print("Windows Instructions\n")

    print("Step 1: Type 'cmd' or 'command prompt' in the search bar.")
    print("Step 2: Right click on the command prompt icon and select 'run in administrator mode'.")
    print("Step 3: Navigate to the USB drive's directory. Type 'E:' Check in explorer if unsure of drive letter.")
    print("Step 4: To run the script, type: python Midtown.py.")
    print("Step 5: Follow the prompt to start or exit the program.")
    print("Step 6: The program will now retrieve the information.")
    print("        If there is already a record of this machine, the program will notify you and ask if you would like to update the record and save the changes.")
    print("        If the information has not changed, the program will notify you and end.")
    print("Step 7: Follow the prompts to save any changes or add a new machine to the csv file.")
    print("Step 8: Once the program has completed, right-click on the thumb drive icon (bottom right) and select 'eject' to safely remove the drive. You can now move on to the next machine.\n")

    print("Trouble Shooting Windows\n")

    print("See below for common problems and suggested solutions for running the script.\n")

    print("Internet connection:")
    print("    Ensure the machine is connected to a network. Some features will fail without an active internet connection, such as the internet speed test.\n")

    print("Dependencies:")
    print("    Ensure the machine is running the latest version of Python.")
    print("    The script will attempt to install any missing python packages. If there is a failure of any kind, the packages can be manually installed from Python using pip.")
    print("    Please see the commands below.\n")

    print("    'pip install pandas'          # Installs Pandas library")
    print("    'pip install speedtest-cli'   # Installs speedtest library")
    print("    'pip install psutil'          # Installs psutil library")
    print("    'pip install getmac'          # Installs getmac library")
    print("    'pip install --upgrade pip'   # Installs pip")
    print("    'pip install py-cpuinfo'      # Installs cpuinfo\n")
                                                                               
    startUp()

def helpLin():
    
    print("User instructions for Midtown.py script\n")
    print("The Midtown.py script is designed to collect information from each computer by running the script from a portable USB thumb drive.")
    print("Follow the instructions below for running the script on Linux.")
    print("The instructions will assume you have loaded both the Midtown_IT.csv and Midtown.py script onto your thumb drive.")
    print("Note this script can be run from within a Python IDE or terminal.")

    print("Linux Instructions\n")
    print("Open a new terminal window.")
    print("This script requires a virtual environment (venv) to run on Linux.")
    print("Commands to create venv from home directory")
    print("Command 1: 'sudo su'")
    print("Command 2: 'python3 -m venv myenv'")
    print("Command 3: 'source myenv/bin/activate'")
    print("To Exit venv, simply type exit")
    print("Using the virtual environment to install these packages is recommended to avoid conflicts with system-wide Python packages.\n")

    print("Running the script")
    print("Step 1: Navigate to the USB drive's directory. Example: Type 'cd /media/user/ThumbDrive'")
    print("Note: If you're unsure of the file path, open your file explorer, go to the USB drive folder,")
    print("right-click on the Midtown python file to view the file path.")
    print("Step 2: To run the script, type: python3 Midtown.py")
    print("Step 3: Follow the prompt to start or exit the program.")
    print("Step 4: The program will now retrieve the information.")
    print("If there is already a record of this machine, the program will notify you and ask if you would like to update the record and save the changes.")
    print("If the information has not changed, the program will notify you and end.")
    print("Step 5: Follow the prompts to save any changes or add a new machine to the CSV file.")
    print("Step 6: Once the program has completed, right-click on the thumb drive icon (bottom right) and select 'eject' to safely remove the drive.")
    print("You can now move on to the next machine.\n")

    print("Trouble Shooting Linux\n")

    print("See below for common problems and suggested solutions for running the script.\n")
    print("Check Internet connection:")
    print("Ensure the machine is connected to a network. Some features will fail without an active internet connection, such as the internet speed test.\n")
    print("For any other problems encountered, please reach out to the developer at AdamB@MidtownAutomation.com \n")
    
    startUp()

#The venvHelp function will automatically print the venv creation commands if the user is not running a venv.
def venvHelp():
    print("This script requires a virtual environment (venv) to run on Linux.")
    print("Commands to create venv from home directory:  ")
    print("Command 1: 'sudo su'")
    print("Command 2: 'python3 -m venv myenv'")
    print("Command 3: 'source myenv/bin/activate'")
    print("Navigate to thumb drive directory and run script")
    print("To Exit venv, simply type exit")
# Get user inputs and ceck os type. Run selected function Windows or Linux or help menu.
def startUp():
    
    operatingsys = input('Enter [1] to start, [2] to exit or [3] for help menu:')
    if operatingsys =='1' and os.name == 'nt':
        pipWin()
        
    elif operatingsys =='1' and os.name != 'nt' and sys.prefix != sys.base_prefix :
        #print("You are in a venv, running Linux script")
        pipLinux()    

    elif operatingsys =='1' and os.name != 'nt' and sys.prefix == sys.base_prefix : #if prefix == base_prefix you are not in a venv
        #print("You are not in a venv, running venv instructions")
        venvHelp()


    elif operatingsys =='2':
        print("Program fnished, no changes recorded.") 
        
    elif operatingsys =='3' and os.name == 'nt':
        helpWin()
        #print("Starting help menu for Windows.") 
          
    elif operatingsys == '3' and os.name != 'nt':
        helpLin()
        #print("Starting Linux help menu" )
        
    else:
        print ('Invalid response')
        startUp()
        
startUp()    




