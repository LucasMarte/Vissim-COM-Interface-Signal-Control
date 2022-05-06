
import win32com.client as com
			
Vissim = com.Dispatch("Vissim.Vissim")

#open vissim in desired project and scenario
arquive_path = r'D:\path_to_your_arquive.ipnx'
Vissim.LoadNet(arquive_path)
Vissim.ScenarioManagement.LoadScenario(1)

#get the semaforical control and its signal groups
semaforical_control = Vissim.Net.SignalControllers.ItemByKey(1)
semaf_group_1 = semaforical_control.SGs.ItemByKey(1)
semaf_group_2 = semaforical_control.SGs.ItemByKey(2)
 
# get the period and the resolution to know how many steps will be needed in the simulation
Period = Vissim.Simulation.AttValue('SimPeriod')
Resolution = Vissim.Simulation.AttValue('SimRes')
total_steps_in_the_simulation = (Period * Resolution)


for i in range(total_steps_in_the_simulation):
    Vissim.Simulation.RunSingleStep()

    #every 30 simulation seconds
    if (i % (30*Resolution) == 0):
        
        #change the signal group's state alternatively every 30 seconds
        if ((i / (30*Resolution)) % 2 == 0):
            semaf_group_1.SetAttValue('SigState', 3)
            semaf_group_2.SetAttValue('SigState', 1)
        else:
            semaf_group_1.SetAttValue('SigState', 1)
            semaf_group_2.SetAttValue('SigState', 3)

