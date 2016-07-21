import xlrd,shutil,time,math,sys,subprocess,xlwt,csv,numpy as np,os
from numpy import *

def GetInputVariables():
    wb_incidentR = xlrd.open_workbook('input_variables_'+COILSIM_version+'.xlsx')
    try:
        sheet_incidentR = wb_incidentR.sheet_by_name(u'incidentR')
    except:
        print 'no sheet named incidentR'
    # incident radiation on inlet tubes
    for i in range(1, 51):
        for j in range(0, 2*nreac+1):
            if j==0:
                x_position_inlet[i-1] = sheet_incidentR.cell(i, j).value
            elif j<23:
                incidentR_inlet[j-1][i-1] = sheet_incidentR.cell(i, j).value
            else:
                ConvRadratio_inlet[j-23][i-1] = sheet_incidentR.cell(i, j).value
    # incident radiation on outlet tubes
    for i in range(52, 102):
        for j in range(0, 2*nreac+1):
            if j==0:
                x_position_outlet[i-52] = sheet_incidentR.cell(i, j).value
            elif j<23:
                incidentR_outlet[j-1][i-52] = sheet_incidentR.cell(i, j).value
            else:
                ConvRadratio_outlet[j-23][i-52] = sheet_incidentR.cell(i, j).value
    
    if CouplingCorrections==True:
        # incident factor of inlet tubes
        for i in range(103, 153):
            for j in range(0, nreac+1):
                if j!=0:
                    InciFactor_inlet[j-1][i-103] = sheet_incidentR.cell(i, j).value
        # incident factor of outlet tubes
        for i in range(154, 204):
            for j in range(0, nreac+1):
                if j!=0:  
                    InciFactor_outlet[j-1][i-154] = sheet_incidentR.cell(i, j).value
        # convective heat transfer coefficient of inlet tubes
        for i in range(205, 255):
            for j in range(0, nreac+1):
                if j!=0: 
                    HTC_inlet[j-1][i-205] = sheet_incidentR.cell(i, j).value
        # convective heat transfer coefficient of outlet tubes
        for i in range(256, 306):
            for j in range(0, nreac+1):
                if j!=0: 
                    HTC_outlet[j-1][i-256] = sheet_incidentR.cell(i, j).value
        # T gas factor of inlet tubes
        for i in range(307, 357):
            for j in range(0, nreac+1):
                if j!=0: 
                    TgFactor_inlet[j-1][i-307] = sheet_incidentR.cell(i, j).value
        # T gas factor of outlet tubes
        for i in range(358, 408):
            for j in range(0, nreac+1):
                if j!=0: 
                    TgFactor_outlet[j-1][i-358] = sheet_incidentR.cell(i, j).value
        # T gas of inlet tubes
        for i in range(409, 459):
            for j in range(0, nreac+1):
                if j!=0:  
                    Tg_inlet[j-1][i-409] = sheet_incidentR.cell(i, j).value
        # T gas of outlet tubes
        for i in range(460, 510):
            for j in range(0, nreac+1):
                if j!=0:  
                    Tg_outlet[j-1][i-460] = sheet_incidentR.cell(i, j).value
        # T wall of inlet tubes
        for i in range(511, 561):
            for j in range(0, nreac+1):
                if j!=0: 
                    Tw_inlet[j-1][i-511] = sheet_incidentR.cell(i, j).value
        # T wall of outlet tubes
        for i in range(562, 612):
            for j in range(0, nreac+1):
                if j!=0:  
                    Tw_outlet[j-1][i-562] = sheet_incidentR.cell(i, j).value

def Generate_walltemp_coefficient(coil_num, point_num, filename):
    # read the workbook with the wall temperature data
    workbook = xlrd.open_workbook(os.path.join(results_dir,filename))
    # used for test 
    # workbook = xlrd.open_workbook('ExternalTemp_it1_ALL.xls')
    try:
        mysheet = workbook.sheet_by_name(u'ExternalWallTemperatures')
    except:
        print 'no sheet named ExternalWallTemperatures'
        return
    # set x for the axial position data from excel file
    x = zeros(point_num)
    # read the axial position value for all the coils
    mark = -1; # used as a mark for inlet and outlet coils division
    for row in range(0, point_num):
        x[row] = mysheet.cell(row+1, 0).value
        if x[row] > 11.42956:
            if mark == -1:
                mark = row
    # set x and y for least square approximation
    x_inlet = zeros(mark)
    y_inlet = zeros(mark)
    x_outlet = zeros(point_num-mark)
    y_outlet = zeros(point_num-mark)
    # read the wall temperature data from excel file
    for col in range(0, coil_num+1):
        # change the axial position data into
        if col == 0:
            for row in range(0, point_num):
                if x[row] <= 9.15:
                    x_inlet[row] = 11.609 - x[row]
                elif x[row] <= 10.16717:
                    x_inlet[row] = 2.459 - 0.95*(x[row]-9.15)/1.01717
                elif x[row] <= 10.26717:
                    x_inlet[row] = 11.67617 - x[row]
                elif x[row] <= 11.42956:
                    x_inlet[row] = 1.409 - 0.74*(math.sin((x[row]-10.26717)/0.74))
                elif x[row] <= 12.59195:
                    x_outlet[row-mark] = 0.74 - 0.74*math.cos((x[row]-11.42956)/0.74) + 0.669
                else:
                    x_outlet[row-mark] = x[row] - 12.59195 + 1.409
        else:
            # divide the data into two columns(inlet and outlet coils)
            for row in range(0, point_num):
                if x[row] <= 11.42956:
                    y_inlet[row] = mysheet.cell(row+1, col).value
                    y_inlet[row] = y_inlet[row] + 273.15
                else:
                    y_outlet[row-mark] = mysheet.cell(row+1, col).value
                    y_outlet[row-mark] = y_outlet[row-mark] + 273.15
            # regression
            coef_inlet[col-1] = polyfit(x_inlet,y_inlet,6)
            coef_outlet[col-1] = polyfit(x_outlet,y_outlet,6)

def calculateCoking():
    global cokeThickness
    naxialcoke=26
    cokingRate = [[0.0 for q in xrange(naxialcoke)] for p in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        if IterationTimestep==1:
            src=os.path.join(templatedir,'Huajin'+nr)
        else:
            src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
        name=src+'\general_info.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if (rownumber <> 0):
                cokingRate[i][rownumber-1]=cora*float(row[14].strip())
                rownumber+=1
            else: rownumber+=1
    cokeThickness=cokeThickness+np.divide(cokingRate,float(cokeDensity*1000/timestepinterval)) #coke thickness in the new iteration step in m
    print 'coke calculation done'

def getInitialCokingThickness():
    #Get cokethickness for all reactors from existing case
    cokeT = [[0.0 for q in xrange(naxial)] for p in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        src=os.path.join(templatedir,'Huajin'+nr)
        name=src+'\general_info.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber<>0:
                cokeT[i][rownumber-1]=float(row[13].strip())
                rownumber+=1
            else:
                rownumber+=1
    return cokeT

def getInitialCokingThicknessV31(filename):
    #Get cokethickness for all reactors from existing case
    cokeT = [[0.0 for q in xrange(naxial)] for p in xrange(nreac)]
    src=os.path.join(templatedir,filename)
    wb_results = xlrd.open_workbook(src)
    try:
        sheet_results = wb_results.sheet_by_name(u'CokeThickness')
    except:
        print 'no sheet named CokeThickness'
    for i in range(0,nreac):
        for row in range(0,naxial):
            cokeT[i][row]=sheet_results.cell(row+1, i+1).value
    return cokeT
    
def getPEinit():
    #obtain P/E for each reactor
    C2H4 = [None for q in xrange(nreac)]
    C3H6 = [None for q in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        src=src=os.path.join(templatedir,'Huajin'+nr)
        name=src+'\yields.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber == 3:
                C2H4[i]=row[2].strip()  
            elif rownumber == 7:
                C3H6[i]=row[2].strip()
            rownumber+=1
    return np.sum(np.multiply(np.divide([float(i) for i in C3H6],[float(i) for i in C2H4]),flowrate[0:(nreac)])/np.sum(flowrate[0:(nreac)]))
    
def getPEinit3(filename):
    src=os.path.join(templatedir,filename)
    wb_results = xlrd.open_workbook(src)
    try:
        sheet_results = wb_results.sheet_by_name(u'Statistics')
    except:
        print 'no sheet named Statistics'
    PE=sheet_results.cell(5, 4).value
    return PE
    
def getPE():
    #obtain P/E for each reactor
    C2H4 = [None for q in xrange(nreac)]
    C3H6 = [None for q in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
        name=src+'\yields.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber == 3:
                C2H4[i]=row[2].strip()  
            elif rownumber == 7:
                C3H6[i]=row[2].strip()
            rownumber+=1
    return np.sum(np.multiply(np.divide([float(i) for i in C3H6],[float(i) for i in C2H4]),flowrate[0:(nreac)])/np.sum(flowrate[0:(nreac)]))
    
def readHeatFluxData(namefile,heatfluxtemplate):
    #Fluent axial positions
    axial = [None]*(nreac)
    #heat flux values at Fluent axial positions
    heatflux = [None]*(nreac)
    #Read Excel file 
    #heat flux values at COILSIM axial positions
    COILSIMheatflux = [None]*(nreac)
    if heatfluxtemplate==True:
        wb = xlrd.open_workbook(os.path.join(datdir,namefile))
    else:
        wb = xlrd.open_workbook(os.path.join(results_dir,namefile))
    #Declare arrays for axial position and heat flux for all nreac coils

    #Open sheet of naphtha feedstocks
    sh = wb.sheet_by_index(0)
    for i in range(0,nreac):
        axial[i] = sh.col_values(3*i)
        length=len(axial[i])
        for j in range(0,length-2):
            axial[i][j]=axial[i][j+2]
        axial[i][length-2]=axial[i][length-3]+1
        axial[i][length-1]=None
        
        #Read heat flux profile and convert to kcal/m2
        heatflux[i] = sh.col_values(3*i+1)
        length=len(axial[i])
        for j in range(0,length-2):
            heatflux[i][j]=heatflux[i][j+2]/1000/4.18400
        heatflux[i][length-2]=heatflux[i][length-3]
        heatflux[i][length-1]=None
    #Generate a plot to check if reading is ok
    '''Plots=Plotting()
    Plots.generatePlot('Test',axial, heatflux,'Axial Position [m]',' Heat flux [W/m2]',results_dir,22)'''       
    #Interpolate heatflux to COILSIM axial positions
    for i in range(0,nreac):
        COILSIMheatflux[i]=np.interp(COILSIMaxial, axial[i], heatflux[i])   
    return COILSIMheatflux

def simulateReactors():
    #Perform reactor simulations and fit CIP per reactor to get right COP------------------------------------------------------------------------------------------------------------
    global naxial
    maxiter=100
    x0=[None]*nreac
    x1=[None]*nreac
    xnew=[None]*nreac
    f0=[None]*nreac
    f1=[1.0]*nreac
    treshold=0.005
    convflagarray=[0]*nreac
    convflag=0
    jj=0
    global COP
    
    while((jj<maxiter) and (convflag==0)):
        for i in range(0,nreac):
            if(jj==0):
                x0[i]=CIP[i]
            elif(jj==1):
                x1[i]=1.01*CIP[i]
            elif (convflagarray[i]==0):
                if (abs(f1[i]-f0[i])>1e-6):
                    if (abs(x1[i]-x0[i])>1e-6):   #if the difference in inlet pressure is too small
                        xnew[i]=x1[i]-f1[i]/((f1[i]-f0[i])/(x1[i]-x0[i]))
                    else:
                        xnew[i]=1.01*x1[i]
                else:
                    xnew[i]=1.01*x1[i]
                if xnew[i]<COPset:
                    x0[i]=x1[i]
                    x1[i]=x0[i]+0.2
                else:
                    x0[i]=x1[i]
                    x1[i]=xnew[i]
                if COP[i]<1.2:
                    x1[i]=5

            
        #Make simulation folders in %appdata%
        #Write exp.txt for all 22 reactors and copy reactor.txt,nafta.i and extra files
        os.chdir(workdir)
        filename=os.path.join(workdir,"Projects\simulation.txt")
        f = open(filename, 'w')
        snsim=str(nreac)
        f.write(snsim+' \n')
        f.write('0 \n')
        src=templatedir
        dst=os.path.join(workdir,'Projects\USC')
        
        if os.path.exists(dst): shutil.rmtree(dst)
        time.sleep(4.0)
        shutil.copytree(src,dst)
#        if kk==1 and jj==0:
#            print 'cokeprofile before reactor simulation written to file'
#            for i in range(0,nreac):
#                for j in range(0,naxial):
#                    sheetCOKE2.write(j,i+1,float(cokeThickness[i][j])) 
#        print 'iteration : '+str(iteration)
        for i in range(0,nreac):
            nr=str(i)
            #Copy all files from template folder
            dst=os.path.join(workdir,'Projects\USC\Huajin'+nr)
            #write cokeprofile
            #print 'Cokethickness profile written to the appropriate files before starting the simulation.'
            filename3=os.path.join(dst,'coke.i')
            g=open(filename3,'w')
            g.write('500'+'    0.1\n')
            g.write('251\n')
            cokeprofile=np.interp(np.linspace(0, COILSIMaxial[len(COILSIMaxial)-1], 252),COILSIMaxial[0:len(COILSIMaxial)-1],cokeThickness[i])
            for k in range(0,len(cokeprofile)-1):
                g.write('0 '+str(cokeprofile[k])+'\n')
            g.close
            #Change exp.txt files
            filename2=dst+"\exp.txt"
            g=open(filename2,'w')
            g.write('1 \n')
            g.write('1 ')
            j=0
            q = ['']*(114)
            for k in range(0,27):
                if(k%9 == 0):
                    q[j]+='\n'
                    g.write(q[j])
                    j=j+1
                elif(k == 26):
                    g.write(q[j])     
                q[j]+=str('%1.4f' % float(heatFluxRatio*COILSIMheatflux[i][k]))
                q[j]+=' '
            j=0
            q = ['']*(114)
            for k in range(0,27):
                
                if(k%9 == 0):
                    q[j]+='\n'
                    g.write(q[j])
                    j=j+1
                elif(k == 26):
                    q[j]+='\n'
                    g.write(q[j])     
                q[j]+=str('%1.3f' % float(flowrate[i]))
                q[j]+=' '    
            g.write(str(dilution)+'\n')
            if(jj==0):
                g.write(str(CIT)+' '+str(x0[i])+'\n')
            else:
                g.write(str(CIT)+' '+str(x1[i])+'\n')
            g.close()
         
            #Adjust simulation.txt  
            f.write('USC\Huajin'+nr+'\n')
        f.close()
        #Change directory and run COILSIM
        if (COILSIM_version=='v2' or COILSIM_version=='v3.1'):
            subprocess.call(['C:\Users\yuzhan\AppData\Roaming\coilsim3d\Coilsim.exe'])
        elif (COILSIM_version=='v3.2' or COILSIM_version=='v3.7'):
            subprocess.call(['C:\ProgramData\coilsim3d\Coilsim.exe'])

        
        #Read COP and compare with COP setpoint to get new CIP
        for i in range(0,nreac):
            nr=str(i)
            src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
            name=src+'\general_info.csv'
            ifile  = open(name, "rb")
            data = list(csv.reader(ifile,delimiter=','))
            COP[i]=float(data[naxial][7].strip())
            #print jj,COP[i]
            if(jj==0):
                f0[i]=COP[i]-COPset
            elif(jj==1):
                f1[i]=COP[i]-COPset
            else:
                if (abs(COP[i]-COPset))<treshold:
                    convflagarray[i]=1
                else:
                    f0[i]=f1[i]
                    f1[i]=COP[i]-COPset
            ifile.close()
        if (np.sum(convflagarray)==nreac):
            convflag=1
        print "        CIP iteration: "+ str(jj)
        jj+=1
        #END CIP optimization loop---------------------------------------------------------------------------------------------------------------------------------------------------
    return x1
    
def getmaxCIP():
    inletPressure = [[None for q in xrange(naxial)] for p in xrange(nreac)]
    maxinletPressure=[0.0 for p in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
        #Retrieve CIP profiles for all reactors
        name=src+'\general_info.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber >0:
                inletPressure[i][rownumber-1]=float(row[7].strip())
                rownumber+=1
            else:
                rownumber+=1
        maxinletPressure[i]=np.max(inletPressure[i])
    return maxinletPressure
    
def getmaxTMT():
    externalT = [[None for q in xrange(naxial)] for p in xrange(nreac)]
    maxTMT=[0.0 for p in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
        #Retrieve external wall temperature profiles for all reactors
        name=src+'\general_info.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber >0:
                externalT[i][rownumber-1]=float(row[5].strip())
                rownumber+=1
            else:
                rownumber+=1
        maxTMT[i]=np.max(externalT[i])
    return maxTMT
    
def getTMT():
    externalT = [[None for q in xrange(naxial)] for p in xrange(nreac)]
    for i in range(0,nreac):
        nr=str(i)
        src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
        #Retrieve external wall temperature profiles for all reactors
        name=src+'\general_info.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber >0:
                externalT[i][rownumber-1]=float(row[5].strip())
                rownumber+=1
            else:
                rownumber+=1
    return externalT

def writeResults(filename):
    global T_fluegas
    global Q_absorb
    #Copy results folders to desktop folder
    #Make folder for results for this iteration (delete if already exists)
    if CoulpedSim==True:
        newpath=os.path.join(results_dir,'timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'_iteration'+str(iterationTMT))
    else:
        newpath=os.path.join(results_dir,'timestep'+str(timestep)+'_PEloop'+str(iterationPE))
    if  os.path.exists(newpath): shutil.rmtree(newpath)
    time.sleep(4.0)
    os.mkdir(newpath)
    os.chdir(newpath)
    #Copy all files from relevant %appdata% projects
    for i in range(0,nreac):
        nr=str(i)
        dst=newpath+"\Huajin"+nr
        src=os.path.join(workdir,'Projects\USC\Huajin'+nr)
        shutil.copytree(src,dst)
    
    externalT = [[None for q in xrange(naxial+1)] for p in xrange(nreac)] 
    processgasT = [[None for q in xrange(naxial+1)] for p in xrange(nreac)]
    C2H4 = [None for q in xrange(nreac)]
    C3H6 = [None for q in xrange(nreac)]
    processgasP = [[None for q in xrange(naxial+1)] for p in xrange(nreac)] 
    cokingRate = [[None for q in xrange(naxial+1)] for p in xrange(nreac)]
    cokeThick =  [[None for q in xrange(naxial+1)] for p in xrange(nreac)]
    HeatFlux =  [[None for q in xrange(naxial+1)] for p in xrange(nreac)]
    PEWeighted = [None for q in xrange(nreac)]
    Massflow = [None for q in xrange(nreac)]
    HeatFluxTotal = [None for q in xrange(nreac)]
    
    for i in range(0,nreac):
        nr=str(i)
        src=os.path.join(newpath,'Huajin'+nr)
        #Retrieve external wall temperature profiles for all reactors
        name=src+'\general_info.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber == 0:
                externalT[i][rownumber]="Reactor nr "+str(i+1)
                processgasT[i][rownumber]="Reactor nr "+str(i+1)
                processgasP[i][rownumber]="Reactor nr "+str(i+1)
                cokingRate[i][rownumber]="Reactor nr "+str(i+1)
                cokeThick[i][rownumber]="Reactor nr "+str(i+1)
                HeatFlux[i][rownumber]="Reactor nr "+str(i+1)
                rownumber+=1
            else:
                externalT[i][rownumber]=row[5]
                processgasT[i][rownumber]=row[3]
                processgasP[i][rownumber]=row[7]
                cokingRate[i][rownumber]=row[14]
                cokeThick[i][rownumber]=row[13]
                HeatFlux[i][rownumber]=row[2]
                rownumber+=1
        #Retrieve flow rate through reactor
        name=src+'\exp.txt'
        f=open(name,'r')
        for j,line in enumerate(f):
            if j==7:
                Massflow[i]=float(line.split()[0])
        f.close()
        #Retrieve heat flux
        name=src+'\\results_summary.txt'
        f=open(name,'r')
        for k,line in enumerate(f):
            if k==13:
                HeatFluxTotal[i]=float(str(line).split()[4])
        #Retrieve yields data for all reactors
        name=src+'\yields.csv'
        ifile  = open(name, "rb")
        data = csv.reader(ifile,delimiter=',')
        x=list(data)  
        rownumber=0
        for row in x:
            if rownumber == 3:
                C2H4[i]=row[2]  
            elif rownumber == 7:
                C3H6[i]=row[2]
            rownumber+=1
    
    #Write temperatures profiles to excel file 
    wbk = xlwt.Workbook()
    
    sheet = wbk.add_sheet("ExternalWallTemperatures", cell_overwrite_ok=True)
    sheet2 = wbk.add_sheet("ProcesgasTemperatures", cell_overwrite_ok=True)
    sheet3 = wbk.add_sheet("Yields", cell_overwrite_ok=True)
    sheet4 = wbk.add_sheet("ProcesgasPressure", cell_overwrite_ok=True)
    sheet5 = wbk.add_sheet("CokingRate", cell_overwrite_ok=True)
    sheet6 = wbk.add_sheet("CokeThickness", cell_overwrite_ok=True)
    sheet7 = wbk.add_sheet("Heat Flux", cell_overwrite_ok=True)
    sheet8 = wbk.add_sheet("Statistics", cell_overwrite_ok=True)
    
    sheet.write(0,0,str("Axial position [m]"))
    sheet2.write(0,0,str("Axial position [m]"))
    sheet4.write(0,0,str("Axial position [m]"))
    sheet5.write(0,0,str("Axial position [m]"))
    sheet6.write(0,0,str("Axial position [m]"))
    sheet7.write(0,0,str("Axial position [m]"))
    
    sheet.write(naxial+3,0,"maxTMT")
    sheet5.write(naxial+3,0,"maxCokingRate")
    sheet7.write(naxial+3,0,"maxHeatFlux")
    
    for j in range(0,naxial+1):
        sheet.write(j+1,0,float(COILSIMaxial[j]))
        sheet2.write(j+1,0,float(COILSIMaxial[j]))
        sheet4.write(j+1,0,float(COILSIMaxial[j]))
        sheet5.write(j+1,0,float(COILSIMaxial[j]))
        sheet6.write(j+1,0,float(COILSIMaxial[j]))
        sheet7.write(j+1,0,float(COILSIMaxial[j]))
    for i in range(0,nreac):
        for j in range(0,naxial+1):
            if j==0:
                sheet.write(j,i+1,str(externalT[i][j]))
                sheet2.write(j,i+1,str(processgasT[i][j]))
                sheet4.write(j,i+1,str(processgasP[i][j]))
                sheet5.write(j,i+1,str(cokingRate[i][j]))
                sheet6.write(j,i+1,str(cokeThick[i][j]))
                sheet7.write(j,i+1,str(HeatFlux[i][j]))
            else:
                sheet.write(j,i+1,float(externalT[i][j]))
                sheet2.write(j,i+1,float(processgasT[i][j]))
                sheet4.write(j,i+1,float(processgasP[i][j]))
                sheet5.write(j,i+1,float(cokingRate[i][j]))
                if COILSIM_version=='v3.1':
                    sheet6.write(j,i+1,float(cokeThickness[i][j-1]))
                else:
                    sheet6.write(j,i+1,float(cokeThick[i][j]))
                sheet7.write(j,i+1,float(HeatFlux[i][j]))
                # convert it to a number
                externalT[i][j]=float(externalT[i][j].strip())
        sheet.write(naxial+3,i+1,float(max(externalT[i][1:])))
        sheet5.write(naxial+3,i+1,float(max(cokingRate[i][1:])))
        sheet7.write(naxial+3,i+1,float(max(HeatFlux[i][1:])))
                    
    sheet3.write(1,0,str("C2H4 [wt%]"))
    sheet3.write(2,0,str("C3H6 [wt%]"))
    sheet3.write(4,0,str("PE ratio [-]"))
    
    for i in range(0,nreac):
        sheet3.write(0,i+1,"Reactor nr "+str(i+1))
        sheet3.write(1,i+1,float(C2H4[i]))
        sheet3.write(2,i+1,float(C3H6[i]))
        sheet3.write(4,i+1,float(C3H6[i])/float(C2H4[i]))
    
    sheet8.write(0,0,'Reactor')
    sheet8.write(1,0,'Mass flow [kg/h]')
    sheet8.write(2,0,'P/E')
    sheet8.write(3,0,'Total heat flux per reactor [kW]')
    sheet8.write(5,0,'Mixing cup average P/E')
    sheet8.write(6,0,'Total heat input to all reactors [kW]')
    sheet8.write(8,0,'Flue gas outlet temperature [K]')
    sheet8.write(9,0,'Total heat absorbed by all reactors [kW]')
    sheet8.write(11,0,'Fuel flow rate scaling ratio')
    sheet8.write(12,0,'Incident radiation scaling ratio')
    
    for i in range(nreac):
        sheet8.write(0,i+4,(i+1))
        sheet8.write(1,i+4,Massflow[i])
        sheet8.write(2,i+4,float(C3H6[i])/float(C2H4[i]))
        sheet8.write(3,i+4,HeatFluxTotal[i])
        PEWeighted[i]=Massflow[i]*float(C3H6[i])/float(C2H4[i])
    
    PEaverage=sum(PEWeighted)/sum(Massflow)
    sheet8.write(5,4,float(PEaverage))
    TotalHeat=sum(HeatFluxTotal)
    sheet8.write(6,4,float(TotalHeat))
    sheet8.write(8,4,float(T_fluegas))
    sheet8.write(9,4,float(Q_absorb))
    sheet8.write(11,4,float(FuelScalingRatio))
    sheet8.write(12,4,float(IncidentScalingRatio))
        
    wbk.save(os.path.join(newpath,filename))
    
    print 'Results summary terminated successfully!'

def writeTMTprofile(filename):
    #Write TMT profiles to excel file 
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet("ExternalWallTemperatures", cell_overwrite_ok=True)

    for j in range(0,nreac+1):
        if j==0:
            sheet.write(0,j,str("Axial position [m]"))
        else:
            sheet.write(0,j,str("Reactor nr "+str(j)))
    for j in range(0,naxial):
        sheet.write(j+1,0,float(COILSIMaxial[j]))
    for i in range(0,nreac):
            for j in range(0,naxial):
                sheet.write(j+1,i+1,float(externalTMT[i][j]))
    # save TMT file
    wbk.save(os.path.join(results_dir,filename))
    print 'New TMT profile written successfully!'

def Generate_heatflux(coil_num,filename):
    # calculate net heat flux based on TMTs and incident radiation
    for i in range(0, coil_num):
        for j in range(0, 50):
            # calculate TMT in Kelvin
            Temp_inlet=coef_inlet[i][0]*x_position_inlet[j]**6.0+coef_inlet[i][1]*x_position_inlet[j]**5.0+coef_inlet[i][2]*x_position_inlet[j]**4.0+coef_inlet[i][3]*x_position_inlet[j]**3.0+coef_inlet[i][4]*x_position_inlet[j]**2.0+coef_inlet[i][5]*x_position_inlet[j]+coef_inlet[i][6]
            Temp_outlet=coef_outlet[i][0]*x_position_outlet[j]**6.0+coef_outlet[i][1]*x_position_outlet[j]**5.0+coef_outlet[i][2]*x_position_outlet[j]**4.0+coef_outlet[i][3]*x_position_outlet[j]**3.0+coef_outlet[i][4]*x_position_outlet[j]**2.0+coef_outlet[i][5]*x_position_outlet[j]+coef_outlet[i][6]
            
            if CouplingCorrections==True:
                # correct the incident radiation
                new_incidentR_inlet=incidentR_inlet[i][j]+(Temp_inlet-Tw_inlet[i][j])*InciFactor_inlet[i][j]
                new_incidentR_outlet=incidentR_outlet[i][j]+(Temp_outlet-Tw_outlet[i][j])*InciFactor_outlet[i][j]
                # correct the gas temperature
                new_Tg_inlet=Tg_inlet[i][j]+(Temp_inlet-Tw_inlet[i][j])*TgFactor_inlet[i][j]
                new_Tg_outlet=Tg_outlet[i][j]+(Temp_outlet-Tw_outlet[i][j])*TgFactor_outlet[i][j]
                # calculate heat flux
                heatflux_inlet[i][j]=WallEmissivity*(new_incidentR_inlet*IncidentScalingRatio-StefanBoltzmann*Temp_inlet**4)+HTC_inlet[i][j]*(new_Tg_inlet-Temp_inlet)
                heatflux_outlet[i][j]=WallEmissivity*(new_incidentR_outlet*IncidentScalingRatio-StefanBoltzmann*Temp_outlet**4)+HTC_outlet[i][j]*(new_Tg_outlet-Temp_outlet)
            else:
                # calculate heat flux
                heatflux_inlet[i][j]=WallEmissivity*(incidentR_inlet[i][j]*IncidentScalingRatio-StefanBoltzmann*Temp_inlet**4)*(1.0+ConvRadratio_inlet[i][j])
                heatflux_outlet[i][j]=WallEmissivity*(incidentR_outlet[i][j]*IncidentScalingRatio-StefanBoltzmann*Temp_outlet**4)*(1.0+ConvRadratio_outlet[i][j])
    
    
    # convert the axial positions
    x_inlet = zeros(50)
    y_inlet = zeros(50)
    x_outlet = zeros(50)
    y_outlet = zeros(50)
    # create a workbook for data written
    writebook = xlwt.Workbook(encoding="utf-8")
    writesheet = writebook.add_sheet('heatflux_profile_innerwall')
    # loop all the coil and value of heatflux
    for col in range(0, coil_num):
        for row in range(0, 50):
            x_inlet[row] = x_position_inlet[row]
            x_outlet[row] = x_position_outlet[row]
            # there are for regions for inlet tube but only two for outlet tube
            # inlet tube: 0.6407-1.409, 1.409-1.509, 1.509-2.459, 2.459-11.609
            if x_inlet[row] <= 1.409:
                x_inlet[row] = 10.26717 + 0.74*math.asin((1.409-x_inlet[row])/0.74)
            elif x_inlet[row] <= 1.509:
                x_inlet[row] = 10.16717 + 1.509 - x_inlet[row]
            elif x_inlet[row] <= 2.459:
                x_inlet[row] = 9.15 + 1.01717*(2.459-x_inlet[row])/0.95
            else:
                x_inlet[row] = 11.609 - x_inlet[row]
            # outlet tube: 0.5407-1.409, 1.409-11.609
            if x_outlet[row] <= 1.409:
                x_outlet[row] = 11.42956 + 0.74*math.acos((0.74-(x_outlet[row]-0.669))/0.74)
            else:
                x_outlet[row] = 12.59195 + x_outlet[row] - 1.409

            y_inlet[row] = heatflux_inlet[col][row]
            y_outlet[row] = heatflux_outlet[col][row]
            # the value of inlet tube external wall heatflux will be changed into internal heat flux(56.6 to 45.0)
            y_inlet[row] = y_inlet[row]*56.6/45.0
            # outlet tube: 0.5407-1.409, 1.409-11.609
            if x_outlet[row] <= 1.409:
                # the value of inlet tube external wall heatflux will be changed into internal heat flux(56.6 to 45.0)
                y_outlet[row] = y_outlet[row]*56.6/45.0
            else:
                # the value of outlet tube external wall heatflux will be changed into internal heat flux(66.6 to 51.0)
                y_outlet[row] = y_outlet[row]*66.6/51.0
        # write the heatflux profile for each tube
        writesheet.write(0, col*3, 'Tube ' + str(col+1))
        writesheet.write(1, col*3, 'Axial Position(m)')
        writesheet.write(1, col*3+1, 'Heat Flux(W/m2)')
        for row in range(0, 50):
            writesheet.write(row+2, col*3, x_inlet[50-1-row])
            writesheet.write(row+2, col*3+1, y_inlet[50-1-row])
            writesheet.write(row+50+2, col*3, x_outlet[row])
            writesheet.write(row+50+2, col*3+1, y_outlet[row])
    writebook.save(os.path.join(results_dir,filename))

def Furnace_estimation(coil_num,filename):
    # estimate heat balance in the furnace and obtain flue gas outlet temperature, absorbed heat by the reactor coils, and corresponding
    global T_fluegas
    global Q_absorb
    # define axial position and internal heat flux
    x_pos = zeros(100)
    y_pos = zeros(100)
    heat_absorbed_coil = zeros(coil_num)
    # open heat flux book
    if FirstTime==True:
        workbook = xlrd.open_workbook(os.path.join(datdir,filename))
    else:
        workbook = xlrd.open_workbook(os.path.join(results_dir,filename))
    try:
        mysheet = workbook.sheet_by_name(u'heatflux_profile_innerwall')
    except:
        print 'no sheet named heatflux_profile_innerwall'
        return
    
    # start calculating the total absorbed heat in all reactor coils (w/m2)
    Q_absorb=0
    for i in range(0, coil_num):
        for j in range(0, 100):
            x_pos[j]=mysheet.cell(j+2, 3*i).value
            y_pos[j]=mysheet.cell(j+2, 3*i+1).value
            if j==0:
                heat_absorbed_coil[i]=heat_absorbed_coil[i]+x_pos[j]*y_pos[j]*3.141592654*0.045
            else:
                # inlet coils
                if x_pos[j-1]<12.59195:
                    heat_absorbed_coil[i]=heat_absorbed_coil[i]+(x_pos[j]-x_pos[j-1])*y_pos[j-1]*3.141592654*0.045
                # outlet coils
                else:
                    heat_absorbed_coil[i]=heat_absorbed_coil[i]+(x_pos[j]-x_pos[j-1])*y_pos[j-1]*3.141592654*0.051
        # the last segment
        heat_absorbed_coil[i]=heat_absorbed_coil[i]+(22.79195-x_pos[99])*y_pos[99]*3.141592654*0.051
        Q_absorb=Q_absorb+heat_absorbed_coil[i]
    # total absorbed heat (kw/m2)
    Q_absorb=Q_absorb/1000.0
    
    # current heat release
    Q_heat_release=Basic_heat_release*FuelScalingRatio
    # percentage of heat loss through furnace refractory
    Q_loss=0.01*Q_heat_release
    # calculate heat taken away by the flue gas
    Q_fluegas=Q_heat_release-Q_loss-Q_absorb
    
    # parameters for enthalpy calculation
    N2_para_high=[2.95257637E+00,1.39690040E-03,-4.92631603E-07,7.86010195E-11,-4.60755204E-15,-9.23948688E+02,5.87188762E+00]
    N2_para_low=[3.53100528E+00,-1.23660988E-04,-5.02999433E-07,2.43530612E-09,-1.40881235E-12,-1.04697628E+03,2.96747038E+00]
    O2_para_high=[3.66096065E+00,6.56365811E-04,-1.41149627E-07,2.05797935E-11,-1.29913436E-15,-1.21597718E+03,3.41536279E+00]
    O2_para_low=[3.78245636E+00,-2.99673416E-03,9.84730201E-06,-9.68129509E-09,3.24372837E-12,-1.06394356E+03,3.65767573E+00]
    CO2_para_high=[4.63651110E+00,2.74145690E-03,-9.95897590E-07,1.60386660E-10,-9.16198570E-15,-4.90249040E+04,-1.93489550E+00]
    CO2_para_low=[2.35681300E+00,8.98412990E-03,-7.12206320E-06,2.45730080E-09,-1.42885480E-13,-4.83719710E+04,9.90090350E+00]
    H2O_para_high=[2.67703890E+00,2.97318160E-03,-7.73768890E-07,9.44335140E-11,-4.26899910E-15,-2.98858940E+04,6.88255000E+00]
    H2O_para_low=[4.19863520E+00,-2.03640170E-03,6.52034160E-06,-5.48792690E-09,1.77196800E-12,-3.02937260E+04,-8.49009010E-01]
    
    b_upper=0.0
    b_lower=0.0
    loop_end=True
    while loop_end==True:
        E_N2=0
        E_O2=0
        E_CO2=0
        E_H2O=0
        for i in range(0,6):
            if T_fluegas>1000:
                if i<5:
                    E_N2=E_N2+N2_para_high[i]*T_fluegas**(i+1)/(i+1)
                    E_O2=E_O2+O2_para_high[i]*T_fluegas**(i+1)/(i+1)
                    E_CO2=E_CO2+CO2_para_high[i]*T_fluegas**(i+1)/(i+1)
                    E_H2O=E_H2O+H2O_para_high[i]*T_fluegas**(i+1)/(i+1)
                else:
                    E_N2=E_N2+N2_para_high[i]
                    E_O2=E_O2+O2_para_high[i]
                    E_CO2=E_CO2+CO2_para_high[i]
                    E_H2O=E_H2O+H2O_para_high[i]
            else:
                if i<5:
                    E_N2=E_N2+N2_para_low[i]*T_fluegas**(i+1)/(i+1)
                    E_O2=E_O2+O2_para_low[i]*T_fluegas**(i+1)/(i+1)
                    E_CO2=E_CO2+CO2_para_low[i]*T_fluegas**(i+1)/(i+1)
                    E_H2O=E_H2O+H2O_para_low[i]*T_fluegas**(i+1)/(i+1)
                else:
                    E_N2=E_N2+N2_para_low[i]
                    E_O2=E_O2+O2_para_low[i]
                    E_CO2=E_CO2+CO2_para_low[i]
                    E_H2O=E_H2O+H2O_para_low[i]
                    
        # enthalpy difference (kJ/kmol) between T_fluegas and reference temperature (298.15 K)
        E_N2=(E_N2*8.3145-0.0)*0.7193
        E_O2=(E_O2*8.3145-0.0)*0.0174
        E_CO2=(E_CO2*8.3145+393510.0)*0.0844
        E_H2O=(E_H2O*8.3145+241826.0)*0.1789
        # total enthalpy per unit molar flow rate
        E_total=E_N2+E_O2+E_CO2+E_H2O
        
        # flue gas molar flow rate (kmol/s)
        Flow_fluegas=Basic_fluegas_flow*FuelScalingRatio/3600
        # enthalpy of the flue gas at temperature of T_fluegas
        Q_fluegas_underT=Flow_fluegas*E_total
        # error
        delta_error=Q_fluegas_underT-Q_fluegas
        
        if abs(delta_error)<0.1:
            break
        if delta_error>0.0:
            b_upper=T_fluegas
            if b_lower==0.0:
                T_fluegas=T_fluegas-10.0
            else:
                T_fluegas=(b_upper+b_lower)/2.0
        if delta_error<0.0:
            b_lower=T_fluegas
            if b_upper==0.0:
                T_fluegas=T_fluegas+10.0
            else:
                T_fluegas=(b_upper+b_lower)/2.0
        if ( (T_fluegas<200) or (T_fluegas>6000) ):
            print 'warning: flue gas temperature T: '+ str(T_fluegas) + '  out of the range'
    #print 'Flue gas outlet temperature:  ' + str(T_fluegas) + ' K'


def IncidentRadiationScalingFactor(caseType,timestep,iterationPE,iterationTMT):
    # calculate the scaling of incident radiation based on the temperature change
    global T_fluegas
    global T_fluegas_base
    global IncidentScalingRatio
    global IncidentLowerLimit
    global IncidentUpperLimit
    # difference of the two flue gas outlet temperatures
    deltaT=T_fluegas-T_fluegas_base
    IncidentScalingRatio_update=1.0+4*deltaT/T_fluegas_base+6*(deltaT/T_fluegas_base)**2+4*(deltaT/T_fluegas_base)**3+(deltaT/T_fluegas_base)**4
    RatioDifference=IncidentScalingRatio_update-IncidentScalingRatio
    #print IncidentScalingRatio_update
    #print 'difference' + str(RatioDifference)
    # incident radiation should be smaller
    if RatioDifference<0.0:
        IncidentUpperLimit=IncidentScalingRatio
        if IncidentLowerLimit==0.0:
            IncidentScalingRatio=IncidentScalingRatio-0.01
        else:
            IncidentScalingRatio=(IncidentLowerLimit+IncidentUpperLimit)/2.0
    # incident radiation should be larger
    else:
        IncidentLowerLimit=IncidentScalingRatio
        if IncidentUpperLimit==0.0:
            IncidentScalingRatio=IncidentScalingRatio+0.01
        else:
            IncidentScalingRatio=(IncidentLowerLimit+IncidentUpperLimit)/2.0
    #print 'IncidentScalingRatio_'+str(caseType)+'_incident_timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'_it'+str(iterationTMT)+':   '+str(IncidentScalingRatio)












#--------------------------Set paths for whole routine--------------------------#

# COILSIM version used " v2, v3.1, v3.2, v3.7 "
COILSIM_version='v3.1'

# data directory
global datdir
datdir='C:\\work\\5 USCfurnace Simulation\\run'
if not os.path.exists(datdir): print('Folder of input does not exist!')

# result directory "data directory+Results"
global results_dir
results_dir = os.path.join(datdir, 'Results')
if not os.path.exists(results_dir): print('Folder of results does not exist!')

# working directory where COILSIM.exe is located
global workdir
if (COILSIM_version=='v2' or COILSIM_version=='v3.1'):
    workdir='C:\\Users\\yuzhan\\AppData\\Roaming\\coilsim3d'
elif (COILSIM_version=='v3.2' or COILSIM_version=='v3.7'):
    workdir='C:\\ProgramData\\coilsim3d'
else:
    print "COILSIM version:  "+COILSIM_version+" does not exist"
    sys.exit()

if not os.path.exists(workdir): print('Working directory does not exist!')

# template directory where the raw files needed for the simulation are stored
global templatedir
templatedir=os.path.join(datdir,'v3.1 template propane') #----------------------------------------change this
print templatedir
if not os.path.exists(templatedir): print('Template directory does not exist!')

# heat flux source
HeatFluxProfileTemplate='HuajinUSC_heatflux_Propane_template.xls' #----------------------------------------change this
#--------------------------Set paths for whole routine--------------------------#


#--------------------------Simulation settings--------------------------#
# constants
cokeDensity=1600 #kg m^-3
StefanBoltzmann=5.670367e-8
nreac=22

# initial operating/geometry conditions
WallEmissivity=0.85
CIT=580
dilution=0.5
COPset=1.76

# TMT under relaxation factor for steaty state simulation
TMTloopalpha=0.5
# Note!!!!!!!!! for single run simulation the heatFluxRatio has to be set to 1
heatFluxRatio=1
# Note!!! for version 3.7 the default value is 1.04, for version 3.1 the value is 1
IncidentScalingRatio=1

# furnace section
FuelScalingRatio=0.95

# temperature of flue gas
global Q_absorb
global T_fluegas
global T_fluegas_base
global IncidentLowerLimit
global IncidentUpperLimit
T_fluegas_base=1368.71875
T_fluegas=1380.0

# heat release from the base case (kW)
Basic_heat_release=14157.0

# flue gas molar flow rate from the base case (kmol/s)
Basic_fluegas_flow=725.821542




# ------ values may need to change ------ #
# flow type
caseType='Propane'    # Options: 'original','coke','COT','PE'

# !!!!!!!!!!!!!!! correction for v3.1 - this is because the coking rate calculation in version 3.1 is not correct!!!!!!!!!!!!!!!!!!
ResultsFilenameV31='Reactor_Results_template.xls'   # name of the result.xls in the template folder

# introduce correction factors at different TMT in the heat flux calculation
CouplingCorrections=False

# the timestep that the simulation starts from
startTimestep=0                               #----------------------------------------change this

# simulation type
RunLengthSim=True     # 'True' perform run length simulation; 'False' no run length simulation will be performed;
OneTimeSimulation=False  # 'True' when the simulation is performed only once with heat flux from external source

# parameters only for run length simulation
CoulpedSim=True       # 'True' coupled run length simulation; 'False' standalone run length simulation;
timestepinterval=50  # the interval of timestep [h]
Maxstep=50               # maximum run length step
cora=0.50               # coking rate scaling factor
TMTmax=1125             # end-of-run criteria TMT
CIPmaximum=5.46         # end-of-run criteria CIP
# ------ values may need to change ------ #

# assign values for different flow types
if caseType=='Propane':
    flowrate=[312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636,312.636] #Propane
if caseType=='original':
    flowrate=[329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090,329.090] #ORIGINAL
if caseType=='original10+':
    flowrate=[361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999,361.999] #ORIGINAL10M
if caseType=='original10-':
    flowrate=[296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181,296.181] #ORIGINAL10L
if caseType=='coke':
    # v3.1 FLOWDIS_coking
    flowrate=[366.201,338.78,327.189,316.456,308.64,307.943,312.185,321.119,331.493,341.821,351.742,366.35,337.416,324.029,314.813,307.877,306.476,311.556,321.048,331.779,343.341,351.726]
    # v2.1 FLOWDIS_coking_converged (it5)
    #flowrate=[365.05,338.016,326.598,316.072,308.425,307.832,312.179,321.10,331.55,342.01,352.004,365.63,337.159,324.052,315.006,308.267,306.993,312.138,321.534,332.331,343.752,352.252]
if caseType=='COT':
    # v3.1 FLOWDIS_COT
    flowrate=[352.764,335.827,327.743,320.917,316.2,315.062,317.582,322.999,330.16,337.655,347.098,351.854,335.031,325.96,319.874,315.001,313.837,316.856,322.756,329.806,338.055,346.943]
    # v2.1 FLOWDIS_COT_Converged (it11)
    #flowrate=[352.16,335.31,327.46,320.80,316.22,315.16,317.52,322.87,330.00,337.68,347.18,351.60,334.81,326.04,320.03,315.43,314.30,317.19,322.91,329.80,338.19,347.32]
if caseType=='PE':
    # v3.1 FLOWDIS_PE
    flowrate=[351.697,336.053,328.25,321.592,316.875,315.636,317.709,322.796,329.838,336.87,346.137,351.079,335.599,326.743,320.651,315.852,314.507,317.121,322.7,329.207,337.199,345.869]
    # v2.1 FLOWDIS_PE_converged (it3)
    #flowrate=[351.990,335.813,327.862,321.059,316.218,315.077,317.350,322.776,329.780,337.236,346.529,351.760,335.553,326.607,320.399,315.553,314.299,317.143,322.938,329.663,337.783,346.592]
if caseType=='TMT':
    # v3.1 FLOWDIS_TMT
    flowrate=[375.682,340.299,326.017,312.098,303.496,303.231,308.321,319.853,334.179,345.319,355.113,376.433,338.29,321.35,310.307,303.058,301.884,307.76,319.406,335.165,347.341,355.378]
if caseType=='Best':
    # v3.1 FLOWDIS_Best
    flowrate=[376.036,337.066,326.974,314.787,305.243,304.359,309.757,320.404,331.231,342.67,356.585,376.185,336.261,323.539,312.753,304.171,302.451,308.971,320.31,331.468,342.191,356.568]
if ((caseType<>'Propane') and (caseType<>'original') and (caseType<>'original10+') and (caseType<>'original10-') and (caseType<>'coke') and (caseType<>'COT') and (caseType<>'PE') and (caseType<>'TMT') and (caseType<>'Best')):
    print 'caseType is not valid, please select one of the following: original, coke, COT, PE'
    sys.exit()
#--------------------------Simulation settings--------------------------#



#--------------------------Simulation initialization--------------------------#
#Read COILSIM axial positions from reactor.txt
os.chdir(datdir)
wb = xlrd.open_workbook('axial_position.xls')
sh = wb.sheet_by_index(0)
COILSIMaxial = sh.col_values(0)
length=len(COILSIMaxial)

# number of values in the first column
naxial=length-1
for j in range(0,length-1):
    COILSIMaxial[j]=COILSIMaxial[j+1]
print COILSIMaxial

# CIP and COP initialization 
CIP=[2.548]*nreac
COP=[0.0]*nreac

# define/initialization incident radiation profile
global x_position_inlet
global x_position_outlet
global incidentR_inlet
global incidentR_outlet
global ConvRadratio_inlet
global ConvRadratio_outlet
x_position_inlet = zeros(50)
x_position_outlet = zeros(50)
incidentR_inlet = zeros((nreac,50))
incidentR_outlet = zeros((nreac,50))
ConvRadratio_inlet = zeros((nreac,50))
ConvRadratio_outlet = zeros((nreac,50))

if CouplingCorrections==True:
    global InciFactor_inlet
    global InciFactor_outlet
    global HTC_inlet
    global HTC_outlet
    global TgFactor_inlet
    global TgFactor_outlet
    global Tg_inlet
    global Tg_outlet
    global Tw_inlet
    global Tw_outlet
    
    InciFactor_inlet = zeros((nreac,50))
    InciFactor_outlet = zeros((nreac,50))
    HTC_inlet = zeros((nreac,50))
    HTC_outlet = zeros((nreac,50))
    TgFactor_inlet = zeros((nreac,50))
    TgFactor_outlet = zeros((nreac,50))
    Tg_inlet = zeros((nreac,50))
    Tg_outlet = zeros((nreac,50))
    Tw_inlet = zeros((nreac,50))
    Tw_outlet = zeros((nreac,50))

# define/initialization coefficients of the polynomial used in TMT 
global coef_inlet
global coef_outlet
coef_inlet = zeros((nreac,7))
coef_outlet = zeros((nreac,7))

# define TMTs in Kelvin and Heat Fluxes in W/m2
global heatflux_inlet
global heatflux_inlet
heatflux_inlet = zeros((nreac,50))
heatflux_outlet = zeros((nreac,50))
#--------------------------Simulation initialization--------------------------#






#--------------------------Start simulation--------------------------#

# read input variables such as incident radiation, flue gas temperature and so on
GetInputVariables()
# read heat flux from template
isheatfluxtemplate=True
COILSIMheatflux=readHeatFluxData(HeatFluxProfileTemplate,isheatfluxtemplate)
# read coke thickness
global cokeThickness
cokeThickness=[[0.0 for q in xrange(naxial)] for p in xrange(nreac)]

if COILSIM_version=='v3.1':
    cokeThickness=getInitialCokingThicknessV31(ResultsFilenameV31)
else:
    cokeThickness=getInitialCokingThickness()



# check if run length simulation is required
if RunLengthSim==True:
    print "start run length simulation"
    IterationTimestep=1
    # get previous P/E as shooting targer for next timestep
    #PEinit=getPEinit()
    PEinit=getPEinit3(ResultsFilenameV31)
    print PEinit
    OneTimeSimulation=False
else:
    print "start steady state simulation"
    IterationTimestep=0
    CoulpedSim=True

#---------------- run length loop ----------------#
RunLengthLoopConv=False
while RunLengthLoopConv==False:
    # calculate time step
    timestep=startTimestep+IterationTimestep*timestepinterval
    # update coke thickness
    if RunLengthSim==True:
        calculateCoking()
        print "-------------Run length timestep: " + str(timestep) + "-------------"
    
    #---------------- P/E loop ----------------#
    PELoopConv=False
    iterationPE=0
    FirstTime=True
    while PELoopConv==False:
        if RunLengthSim==True:
            print "P/E iteration: " + str(iterationPE)
        
        # coupled run length simulation, TMT loop is needed
        if CoulpedSim==True:
            #---------------- TMT loop ----------------#
            TMTLoopConv=False
            iterationTMT=0
            TMTloopCount=0
            while TMTLoopConv==False:
                print "    TMT iteration: " + str(iterationTMT)
                CIP=simulateReactors()
                print CIP
                # calculate the furnace section
                if FirstTime==True:
                    Furnace_estimation(nreac,HeatFluxProfileTemplate)
                    FirstTime=False
                else:
                    Furnace_estimation(nreac,HeatFluxProfileSource)
                writeResults('Reactor_Results_'+str(caseType)+'_timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'_it'+str(iterationTMT)+'.xls')
                # first iteration (one more simulation is necessary)
                if iterationTMT==0:
                    externalTMT=getTMT()
                # not the first iteration, compare the old TMT and new TMT profile
                else:
                    externalTMT_old=externalTMT
                    externalTMT=getTMT()
                    TMTerror=abs(array(externalTMT)-array(externalTMT_old))
                    TMTerrormax=0
                    for i in range(0,nreac):
                        if TMTerrormax<np.max(TMTerror[i]):
                            TMTerrormax=np.max(TMTerror[i])
                    # test how many times TMT is lower than a certain value
                    if (TMTerrormax<2.5 or np.mean(TMTerror)<0.25):
                        print np.mean(TMTerror)
                        TMTloopCount+=1
                    # TMT is converged
                    if TMTerrormax<0.5:
                        print "Maximum TMT error:  " + str(TMTerrormax)
                        print "TMT loop finished (reached TMT error)"
                        break
                    if TMTloopCount==50:
                        print "Maximum TMT error:  " + str(TMTerrormax)
                        print "TMT loop finished (reached maximum iteration)"
                        break
                    # TMT loop is still needed
                    else:
                        # update TMT profile
                        externalTMT=array(externalTMT)*TMTloopalpha+array(externalTMT_old)*(1.0-TMTloopalpha)
                # when the simulation is performed only once with heat flux from external source
                if OneTimeSimulation==True:
                    break
                # write new TMT profile to a file
                writeTMTprofile('TMTprofile_'+str(caseType)+'_timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'_it'+str(iterationTMT+1)+'.xls')
                # calculate coefficients of the polynomial used in TMT
                Generate_walltemp_coefficient(nreac,naxial,'TMTprofile_'+str(caseType)+'_timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'_it'+str(iterationTMT+1)+'.xls')
                
                # furnace calculation
                FurnaceHeatBalance=False
                T_fluegas_old=0.0
                IncidentUpperLimit=0.0
                IncidentLowerLimit=0.0
                while FurnaceHeatBalance==False:
                    # calculate heat flux profile
                    HeatFluxProfileSource='HuajinUSC_heatflux_'+str(caseType)+'_incident_timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'_it'+str(iterationTMT+1)+'.xls'
                    Generate_heatflux(nreac,HeatFluxProfileSource)
                    # calculate the furnace section
                    Furnace_estimation(nreac,HeatFluxProfileSource)
                    # check if the furnace heat balance is converged
                    BalanceError=abs(T_fluegas-T_fluegas_old)
                    if BalanceError<0.5:
                        print 'T_flue:  ' + str(T_fluegas)
                        print 'Incident Radiation:  ' + str(IncidentScalingRatio)
                        break
                    else:
                        T_fluegas_old=T_fluegas
                    # calculate incident radiation scaling factor
                    IncidentRadiationScalingFactor(caseType,timestep,iterationPE,iterationTMT+1)
                
                # read heat flux from previous results
                isheatfluxtemplate=False
                COILSIMheatflux=readHeatFluxData(HeatFluxProfileSource,isheatfluxtemplate)
                if iterationTMT<>0:
                    print "Maximum TMT error:  " + str(TMTerrormax)
                iterationTMT+=1
            #---------------- TMT loop end ----------------#
        
        # standalone run length simulation
        else:
            CIP=simulateReactors()
            print CIP
            T_fluegas=0.0
            Q_absorb=0.0
            FuelScalingRatio=0.0
            IncidentScalingRatio=0.0
            writeResults('Reactor_Results_'+str(caseType)+'_timestep'+str(timestep)+'_PEloop'+str(iterationPE)+'.xls')
        
        # calculate mixing-cup P/E
        PEcurrent=getPE()
        print 'PEWeighted: '+str(PEcurrent)
        # P/E loop is not needed if there is no run length simulation to be performed
        if RunLengthSim==False:
            print "simulation is completed"
            break
        # P/E is converged
        elif ((abs(PEcurrent-PEinit)/PEinit))<0.0005:
            print "P/E loop finished"
            break
        # shoot on P/E value
        else:
            # first iteration, set a fixed value of scaling ratio
            if iterationPE==0:
                # coupled simulation
                #if CoulpedSim==True:
                #    IncidentScalingRatio_old=IncidentScalingRatio
                #    IncidentScalingRatio=IncidentScalingRatio+0.01
                #    print "Incident radiation ratio:  " + str(IncidentScalingRatio)
                if CoulpedSim==True:
                    FuelScalingRatio_old=FuelScalingRatio
                    FuelScalingRatio=FuelScalingRatio+0.01
                    print "Fuel flow rate scaling ratio:  " + str(FuelScalingRatio)
                # uncoupled simulation
                else:
                    heatFluxRatio_old=heatFluxRatio
                    heatFluxRatio=heatFluxRatio+0.01
                    print "Heat flux ratio:  " + str(heatFluxRatio)
            else:
                # coupled simulation
                #if CoulpedSim==True:
                #    IncidentScalingRatio_new=IncidentScalingRatio+(IncidentScalingRatio_old-IncidentScalingRatio)/(PEold-PEcurrent)*(PEinit-PEcurrent)
                #    IncidentScalingRatio_old=IncidentScalingRatio
                #    IncidentScalingRatio=IncidentScalingRatio_new
                #    print "Incident radiation ratio:  " + str(IncidentScalingRatio)
                if CoulpedSim==True:
                    FuelScalingRatio_new=FuelScalingRatio+(FuelScalingRatio_old-FuelScalingRatio)/(PEold-PEcurrent)*(PEinit-PEcurrent)
                    FuelScalingRatio_old=FuelScalingRatio
                    FuelScalingRatio=FuelScalingRatio_new
                    print "Fuel flow rate scaling ratio:  " + str(FuelScalingRatio)
                # uncoupled simulation
                else:
                    heatFluxRatio_new=heatFluxRatio+(heatFluxRatio_old-heatFluxRatio)/(PEold-PEcurrent)*(PEinit-PEcurrent)
                    heatFluxRatio_old=heatFluxRatio
                    heatFluxRatio=heatFluxRatio_new
                    print "Heat flux ratio:  " + str(heatFluxRatio)
            # store mixing-cup P/E
            PEold=PEcurrent
            
        iterationPE+=1
    #---------------- P/E loop end ----------------#
    
    # maximum TMT value
    maxTMTvalue=getmaxTMT()
    print 'timestep: '+str(timestep)+'   maxTMT: '+str(np.max(maxTMTvalue))+'\n'
    # maximum CIP value
    maxCIPvalue=getmaxCIP()
    print 'timestep: '+str(timestep)+'   maxCIP: '+str(np.max(maxCIPvalue))+'\n'
    
    # simulation is completed if there is no run length simulation to be performed
    if RunLengthSim==False:
        break
    # run length simulation stopped (TMT criterion)
    if np.max(maxTMTvalue)>=TMTmax:
        print "run length simulation finished (reached Maximum TMT)"
        break
    # run length simulation stopped (P/E criterion)
    if np.max(maxCIPvalue)>=CIPmaximum:
        print "run length simulation finished (reached Maximum CIP)"
        break
    # run length simulation finished
    if IterationTimestep==Maxstep:
        print "run length simulation finished (reached maximum iteration)"
        break
    IterationTimestep+=1
#---------------- run length loop end ----------------#
#--------------------------End simulation--------------------------#

print ""