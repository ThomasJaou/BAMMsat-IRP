import math as mp 
import matplotlib.pyplot as plt
import csv

"Gas constant"
n = 1.4
k = 1.4 
Z = 0.97 #Gas compressibility factor 
R = 287.05 #Air constant Ratio 


"Pressure Vessel constant"
V = (10e-2)**3  #Volume of the Presure Vessel (10e-2)**3 
d = 2e-3 #diameters of the hole 
A =  mp.pi * (d/2)**2 # Area of the hole 
Cd = 1 #discharge coeff (might be changed)


"Environment constant"
P_initial = 1e5 #Initial pressure inside the vessel  
P_outside = 1e5 #PRessure outside the vessel 
Ti = 293 #initial temperature inside the vessel
rhoi = P_initial/(R*Ti) #initial density inside the vessel

Q_prod = 7.3e-3 #Mass flow rate of the coolant gas, put 0 if none 

"Defined Constant"

C1 = (k+1)/(k-1) #power
C2 = mp.sqrt((2/(k+1))**C1) #square root     
C3 =Cd*A*mp.sqrt(Z*k*R)/V

def Vessel_decompression(P_initial, P_outside):
    
    #intial condition 
    dt = 0.01
    t  = [0]
    
    tempo = (2/(k+1))**(k/(k-1))
    P_limite =  P_outside/tempo 
    
    P= [P_initial]
    rho = [rhoi] 
    T = [Ti]
   
    #3/dt was chosen in order to have all the graph on the same timescale    
    while len(t) < 3/dt  :
    
        if P[-1] >= P_limite :
            
            #calculation of new rho
            rho_i = rho[-1]-(rho[-1]*C2*C3*mp.sqrt(T[-1]) - Q_prod/V)*dt
            rho.append(rho_i)
            
            #calculation of new P 
            P_i = P[-1]*(rho[-1]/rho[-2])**n
            P.append(P_i)
            
            #calculation of new T 
            T_i = P[-1]/(rho[-1]*R)
            T.append(T_i)
            
            t.append(t[-1]+dt) 
            
        
        else : 
            Mt = mp.sqrt( 2/(k-1)*( (P[-1]/P_outside)**((k-1)/k)   -1  )   )  
            
            C3bis = C3*Mt 
            C2bis = mp.sqrt((1+((k-1)/2)*Mt**2)**(-C1) )
            
            
            #calculation of new rho
            rho_i = rho[-1]-(rho[-1]*C2bis*C3bis*mp.sqrt(T[-1])- Q_prod/V)*dt
            rho.append(rho_i)
            
            #calculation of new P 
            P_i = P[-1]*(rho[-1]/rho[-2])**n
            P.append(max(P_i, P_outside))
            
            #calculation of new T 
            T_i = P[-1]/(rho[-1]*R)
            T.append(T_i)
            
            
            t.append(t[-1]+dt)
        
    #print(rho)
    plt.plot(t, P)
    plt.xlabel("Time")
    plt.ylabel("Pressure")
    plt.ticklabel_format(style='sci', axis='y', scilimits=(0,0))
    
    #is used to produce a csv file, can be removed 
    folder = "coolant_produced/coolant_d{}.csv".format(d*1000)
    with open(folder, mode='w', newline='') as fichier_csv:
        writer = csv.writer(fichier_csv)
        # Écrire l'en-tête (facultatif, mais utile pour référence)
        writer.writerow(['x_value', 'y_value'])
        # Écrire les données ligne par ligne
        for x, y in zip(t, P):
            writer.writerow([x, y])
    
    return()

Vessel_decompression(P_initial, P_outside)