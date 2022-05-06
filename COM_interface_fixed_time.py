
import win32com.client as com
			
Vissim = com.Dispatch("Vissim.Vissim")

arquive_path = r'D:\gugão\A_deFacúl\! estágios e coisas de trabalho\lab4c\lab4c.inpx'
Vissim.LoadNet(arquive_path)
Vissim.ScenarioManagement.LoadScenario(1)

semaforical_control = Vissim.Net.SignalControllers.ItemByKey(1)
semaf_group_1 = semaforical_control.SGs.ItemByKey(1)
semaf_group_2 = semaforical_control.SGs.ItemByKey(2)
 
Period_of_the_simulation = Vissim.Simulation.AttValue('SimPeriod')
Resolution = Vissim.Simulation.AttValue('SimRes')
quantidade_total_de_passos = (Period_of_the_simulation* Resolution)

for i in range(quantidade_total_de_passos):
    Vissim.Simulation.RunSingleStep()

    if (i % (30*Resolution) == 0):
        print(i)

        if ((i / (30*Resolution)) % 2 == 0):
            semaf_group_1.SetAttValue('SigState', 3)
            semaf_group_2.SetAttValue('SigState', 1)
        else:
            semaf_group_1.SetAttValue('SigState', 1)
            semaf_group_2.SetAttValue('SigState', 3)


'''
        print(tempo_passado)
        if (estado == 1): estado = 2
        if (estado == 2): estado = 1


Vissim.Simulation.RunSingleStep()
grupo_semaforico_1.SetAttValue('SigState', 3)
grupo_semaforico_2.SetAttValue('SigState', 1)
Vissim.Simulation.RunContinuous()

Vissim.Net.Links.Lanes.ItemByKey(1)
_ = input("Aperte 'Enter' para fechar")
'''
'''
# código de como realizar controle semafórico pelo python COM help. O segundo comentário é realmente útil

# Set the state of a signal controller:
# Note: Once a state of a signal group is set, the attribute "ContrByCOM" is automatically set to True. Meaning the signal group will keep this state until another state is set by COM or the end of the simulation
# To switch back to the defined signal controller, set the attribute signal "ContrByCOM" to False (example see below).
SC_number = 1 # SC = SignalController
SG_number = 1 # SG = SignalGroup
SignalController = Vissim.Net.SignalControllers.ItemByKey(SC_number)
SignalGroup = SignalController.SGs.ItemByKey(SG_number)
new_state = "GREEN" # possible values 'GREEN', 'RED', 'AMBER', 'REDAMBER' and more, see COM Help: SignalizationState Enumeration
SignalGroup.SetAttValue("SigState", new_state)
# Note: The signal controller can only be called at whole simulation seconds, so the state will be set in Vissim at the next whole simulation second, here 199s
# Simulate so that the new state is active in the Vissim simulation:
Sim_break_at = 200 # simulation second [s]
Vissim.Simulation.SetAttValue("SimBreakAt", Sim_break_at)
Vissim.Simulation.RunContinuous() # start the simulation until SimBreakAt (200s)
# Give the control back:
SignalGroup.SetAttValue("ContrByCOM", False)
'''