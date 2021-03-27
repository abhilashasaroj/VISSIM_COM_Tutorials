

## STEP 1: Import Libraries needed to run the code
from __future__ import print_function
import os
# COM-Server
import win32com.client as com

## STEP 2: Connecting the COM Server => Open a new Vissim Window:
Vissim = com.gencache.EnsureDispatch("Vissim.Vissim") #


### STEP 3: Get address of the directory where Vissim file is.
### This is not a necessary step. It is used to let the script know where to find the file you want to run
### The address to folder from where this script is running is stored in Path_of_COM_Basic_Commands_network
Path_of_COM_Basic_Commands_network = os.getcwd()

## You can print out and check if the command os.getwd(), which means - operating system.get working directory gives you the correct address
print (os.getcwd())

## STEP 3: Add name of the Vissim file to the address we fetched in step 3. Store complete address in Filename
Filename = os.path.join(Path_of_COM_Basic_Commands_network, 'sample_network1.inpx')

## STEP 4: Load Vissim inp file on Vissim instance using function Vissim.LoadNet(address to file)
Vissim.LoadNet(Filename)

## STEP 5: Load Vissim layout file on Vissim instance using function Vissim.LoadLayout(address to file)
Filelayout = os.path.join(Path_of_COM_Basic_Commands_network, 'sample_network1.layx')
Vissim.LoadLayout(Filelayout)

## STEP 6: Select random seed for the simulation
Random_Seed = [1]

vols_eastound =[1000,2000,4500]
## for loop to over various random seeds, in our case, it's just one random see which is 21 ##
for r in range(0, len(Random_Seed)):

    ## STEP 7: Set Simulation Run Attriutes
    ## Set Simulation time. Vissim.Simulation.SetAttValue can e used to access and set several Vissim Simulation Run attriutes
    simtime = 900

    ## Set simulation time
    Vissim.Simulation.SetAttValue('SimPeriod', simtime)

    ## Set simulation to use max simulation speed
    Vissim.Simulation.SetAttValue('UseMaxSimSpeed', True)

    ## Set random seed for the simulation run
    Vissim.Simulation.SetAttValue('RandSeed', Random_Seed[r])


    ##STEP 8: 
    #####----- To get access to vehicle input (volumes on link), first we will get vehicle input ids assigned to the links-----####
    ########### Get Vehicle Input ID for each entry links ######################################
    ## Our test model has 4 entry links. 
    ## Doule click vehicle input icon on left sidear of Vissim interface to get Vehicle input id list  
    east_link = 1
    west_link = 3
    north_link = 2
    south_link = 4

    ##STEP 9:
    ###------- Let's access the vehicle inputs for links and create time interval slots to store voume per slot
    ############### SET TIME INTERVAL COLLECTION FOR VOLUME INPUTS ##########################################
    # Get total numer of slots to e added and create intervals
    for timeInt in range(2, (int(simtime/300)+1)):
        # Add timeinterval y accessing vissim timeinterval for vehicle input parameter
        Vissim.Net.TimeIntervalSets.ItemByKey(1).TimeInts.AddTimeInterval(timeInt)
        # Set start  time for the time interval created in aove command
        TimeIntNoNew1 = Vissim.Net.TimeIntervalSets.ItemByKey(1).TimeInts.ItemByKey(timeInt)
        TimeIntNoNew1.SetAttValue('Start',300*(timeInt-1))
        #Set continuous property to e false 
        Vissim.Net.VehicleInputs.ItemByKey(east_link).SetAttValue('Cont('+str(timeInt)+')', False)
        Vissim.Net.VehicleInputs.ItemByKey(west_link).SetAttValue('Cont('+str(timeInt)+')', False)
        Vissim.Net.VehicleInputs.ItemByKey(north_link).SetAttValue('Cont('+str(timeInt)+')', False)
        Vissim.Net.VehicleInputs.ItemByKey(south_link).SetAttValue('Cont('+str(timeInt)+')', False)
        
        
    #######################  SET REAL TIME START TIMES FOR THE SIMULATION ######################## 

    #Get Simulation Resolution attriute from simulation model set in Vissim interfacce
    simRes = Vissim.Simulation.AttValue('SimRes')

    #initiate variale i that represents simulation step 
    i = 0

    # create a while loop to run simulation that loops over variale i.
    while (i<=((simtime-1)*simRes)):
            
        ## set vissim volumes every 5 minute.5 minutes = 300 seconds = 3000 sim steps (the resolution is 10)
        ## Creating if loop that lets volume inputs accessed only when simulation step is a multiple of 3000 sim steps..
        ## Simply put, it lets volume inputs accessed every 5 minutes. 
        ## We will  need to change this value if we change the interval lengths.
        if (i%3000==0):
            vol_int_number = int(i/3000)+1 #Volume interval number = Vehicle Input Number
            print (vol_int_number)
            
            ## Generate name of volume interval to access interval numer for which volume value is eing updated.      
            volume_interval = 'Volume('+str(vol_int_number)+')'
            print (volume_interval)

            # Add the new vehicle inputs. Th
            # This function Vissim.Net.VehicleInputs.ItemByKey(east_link).SetAttValue takes two parameters
            # 1. name of interval eing accessed and 2. volume to e set for this interval
            Vissim.Net.VehicleInputs.ItemByKey(east_link).SetAttValue(volume_interval, int(vols_eastound[vol_int_number-1])*12)
            Vissim.Net.VehicleInputs.ItemByKey(west_link).SetAttValue(volume_interval, int(vols_eastound[vol_int_number-1])*12)
            Vissim.Net.VehicleInputs.ItemByKey(north_link).SetAttValue(volume_interval, int(vols_eastound[vol_int_number-1])*12)
            Vissim.Net.VehicleInputs.ItemByKey(south_link).SetAttValue(volume_interval, int(vols_eastound[vol_int_number-1])*12)

        #The if loop aove is not needed if we don't want to change volume during the run
        ## Run Simulation Step and Increment Value of i
        Vissim.Simulation.RunSingleStep()
        i=i+1
