

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

vols_eastound =[1000,2000,4500, 1000, 2000, 4500, 1000, 2000, 4500, 1000, 2000, 4500]

## for loop to over various random seeds, in our case, it's just one random see which is 21 ##
for r in range(0, len(Random_Seed)):

    ## STEP 7: Set Simulation Run Attriutes
    ## Set Simulation time. Vissim.Simulation.SetAttValue can e used to access and set several Vissim Simulation Run attriutes
    simtime = 3600

    ## Set simulation time
    Vissim.Simulation.SetAttValue('SimPeriod', simtime)
    ## Set simulation to use max simulation speed
    Vissim.Simulation.SetAttValue('UseMaxSimSpeed', True)
    ## Set random seed for the simulation run
    Vissim.Simulation.SetAttValue('RandSeed', Random_Seed[r])

    ##STEP 8: 
    #####----- To get access to signal heads, first we will get vehicle input ids assigned to the links-----####
    ########### Get acess to signal head objects ######################################
    ## Our test model has 4 entry links. 
    obj_ebtr=Vissim.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1)

    #Assigning signal head for phase 6
    obj_wbtr=Vissim.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(1)
    obj_sbtr=Vissim.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(2)
    obj_nbtr=Vissim.Net.SignalControllers.ItemByKey(1).SGs.ItemByKey(2)

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

        # print ('SIM SECOND BEFORE  '+str(Vissim.Simulation.SimulationSecond))
        Vissim.Simulation.RunSingleStep()
        # print ('SIM SECOND AFTER  '+str(Vissim.Simulation.SimulationSecond))
        print (str(i)+'  RUN STEP')

        ## Set start signal states for each signal head. For simplicity
        if (i==0):
                obj_ebtr.SetAttValue("SigState","GREEN")
                obj_wbtr.SetAttValue("SigState","GREEN")
                obj_nbtr.SetAttValue("SigState","RED")
                obj_sbtr.SetAttValue("SigState","RED")

        if ((i<18000) and (i!=0)): 

            if (i%100==40):
                obj_ebtr.SetAttValue("SigState","AMBER")
                obj_wbtr.SetAttValue("SigState","AMBER")
                obj_nbtr.SetAttValue("SigState","RED")
                obj_sbtr.SetAttValue("SigState","RED")

            if (i%100==43):
                obj_ebtr.SetAttValue("SigState","RED")
                obj_wbtr.SetAttValue("SigState","RED") 
                obj_nbtr.SetAttValue("SigState","RED")
                obj_sbtr.SetAttValue("SigState","RED")

            if (i%100==49):
                obj_ebtr.SetAttValue("SigState","RED")
                obj_wbtr.SetAttValue("SigState","RED")
                obj_nbtr.SetAttValue("SigState","GREEN")
                obj_sbtr.SetAttValue("SigState","GREEN")

            if (i%100==92):
                obj_ebtr.SetAttValue("SigState","RED")
                obj_wbtr.SetAttValue("SigState","RED") 
                obj_nbtr.SetAttValue("SigState","AMBER")
                obj_sbtr.SetAttValue("SigState","AMBER")

            if (i%100==96):
                obj_ebtr.SetAttValue("SigState","RED")
                obj_wbtr.SetAttValue("SigState","RED") 
                obj_nbtr.SetAttValue("SigState","RED")
                obj_sbtr.SetAttValue("SigState","RED")


            if (i%100==0):
                obj_ebtr.SetAttValue("SigState","GREEN")
                obj_wbtr.SetAttValue("SigState","GREEN") 
                obj_nbtr.SetAttValue("SigState","RED")
                obj_sbtr.SetAttValue("SigState","RED")


        else:
            obj_ebtr.SetAttValue("ContrByCOM", False)
            obj_wbtr.SetAttValue("ContrByCOM", False)
            obj_nbtr.SetAttValue("ContrByCOM", False)
            obj_sbtr.SetAttValue("ContrByCOM", False)
        
        ## Run Simulation Step and Increment Value of i
        i=i+1


        # Print Signal States
        State_of_Ebtr =  obj_ebtr.AttValue('SigState')
        State_of_Wbtr =  obj_wbtr.AttValue('SigState')
        State_of_Nbtr =  obj_nbtr.AttValue('SigState')
        State_of_Sbtr =  obj_sbtr.AttValue('SigState')
        print (Vissim.Simulation.SimulationSecond)

        ebtr_number =1
        wbtr_number =1
        nbtr_number =2
        sbtr_number =2

        print ('Actual state of SignalHead(%d) is: %s' % (ebtr_number,State_of_Ebtr))
        print ('Actual state of SignalHead(%d) is: %s' % (wbtr_number,State_of_Wbtr))
        print ('Actual state of SignalHead(%d) is: %s' % (nbtr_number,State_of_Nbtr))
        print ('Actual state of SignalHead(%d) is: %s' % (sbtr_number,State_of_Sbtr))

