import xlrd,shutil,time,math,sys,subprocess,xlwt,csv,numpy as np,os
from numpy import *


# this function read the incident radiative heat flux (W/m^2) from the input file
def GetInputVariables():
    rb_IncidentR = xlrd.open_workbook(FileIncidentR)
    sheet_IncidentR = rb_IncidentR.sheet_by_name(u'incidentR')
    # axial position and incident radiation on inlet legs
    for i in range(0,N_furnace_points):
        x_position_inlet[i] = sheet_IncidentR.cell(i+1, 0).value
        for j in range(0,N_reactor):
            IncidentR_inlet[j][i] = sheet_IncidentR.cell(i+1, j+1).value
            ConvRadratio_inlet[j][i] = sheet_IncidentR.cell(i+1, N_reactor+j+1).value
    # axial position and incident radiation on outlet legs
    for i in range(0,N_furnace_points):
        x_position_outlet[i] = sheet_IncidentR.cell((N_furnace_points+1)+i+1, 0).value
        for j in range(0,N_reactor):
            IncidentR_outlet[j][i] = sheet_IncidentR.cell((N_furnace_points+1)+i+1, j+1).value
            ConvRadratio_outlet[j][i] = sheet_IncidentR.cell((N_furnace_points+1)+i+1, N_reactor+j+1).value




# this function read the coke thickness (m) from the template folder
def GetInitialValues():
    rb_ResultsTemplate = xlrd.open_workbook(TempDir+'\\'+FileReactor)
    sheet_CokeThickness = rb_ResultsTemplate.sheet_by_name(u'CokeThickness')
    # axial position and coke thickness
    for i in range(0,N_reactor_axial):
        Coilsim_Axial[i] = sheet_CokeThickness.cell(i+1, 0).value
        for j in range(0,N_reactor):
            CokeThickness[j][i] = sheet_CokeThickness.cell(i+1, j+1).value
    # initial CIP
    sheet_CIP = rb_ResultsTemplate.sheet_by_name(u'ProcesgasPressure')
    for i in range(N_reactor):
        CIP[i] = sheet_CIP.cell(1, i+1).value
    # get values (maximum TMT, maximum CIP, mixing-cup P/E and coking rate) required by run length simulations
    global MaxTMT
    global MaxCIP
    global MixingCupPEtarget
    # get initial mixing-cup P/E
    if RunLengthSim==1:
        sheet_PE = rb_ResultsTemplate.sheet_by_name(u'Statistics')
        MixingCupPEtarget = sheet_PE.cell(5, 4).value
    # get the TMT and its maximum
    sheet_TMT = rb_ResultsTemplate.sheet_by_name(u'ExternalWallTemperatures')
    for i in range(0,N_reactor_axial):
        for j in range(0,N_reactor):
            TMT[j][i]=sheet_TMT.cell(i+1, j+1).value
    # get the maximum TMT
    MaxTMT = np.max(TMT)
    # get the maximum CIP
    MaxCIP = max(CIP)
    # get coking rate
    sheet_CokingRate = rb_ResultsTemplate.sheet_by_name(u'CokingRate')
    for i in range(0,N_reactor_axial):
        for j in range(0,N_reactor):
            CokingRate[j][i] = sheet_CokingRate.cell(i+1, j+1).value




# this function read the heat flux (W/m^2) from the template folder and convert it to (kcal/m^2)
def GetInitialHeatFlux():
    rb_HeatFlux = xlrd.open_workbook(FileHeatFlux)
    sheet_HeatFlux = rb_HeatFlux.sheet_by_name(u'heatflux_profile_innerwall')
    # axial position and heat flux
    HeatFluxTemp=zeros((N_reactor,2*N_furnace_points))
    for i in range(0,2*N_furnace_points):
        Furnace_Axial[i] = sheet_HeatFlux.cell(i+2, 0).value
        for j in range(0,N_reactor):
            Heatflux_Furnace[j][i] = sheet_HeatFlux.cell(i+2, 3*j+1).value
            # convert heat flux to (kcal/m2)
            HeatFluxTemp[j][i] = Heatflux_Furnace[j][i]/1000.0/4.18400
    # interpolate heatflux to COILSIM axial positions
    for i in range(0,N_reactor):
        HeatFlux[i]=np.interp(Coilsim_Axial, Furnace_Axial, HeatFluxTemp[i])




# this function prepare the content to be written to the exp.txt file
def PrepareContent(variable):
    # prepare the flow rate input
    if variable == 'FlowRate':
        for i in range(0,N_reactor):
            # reinitialize flow rate content
            flowrate_content[i]=''
            for j in range(0,N_reactor_axial):
                if FlowRate[i]>100:
                    flowrate_content[i] += str('%1.3f' % float(FlowRate[i]))+' '
                elif FlowRate[i]>10:
                    flowrate_content[i] += str('%1.4f' % float(FlowRate[i]))+' '
                else:
                    flowrate_content[i] += str('%1.5f' % float(FlowRate[i]))+' '
                # every line has 9 values
                if (j+1)%9==0:
                    flowrate_content[i] += '\n'
    # prepare the heat flux input
    elif variable == 'HeatFlux':
        for i in range(0,N_reactor):
            # reinitialize heat flux content
            heatflux_content[i]=''
            for j in range(0,N_reactor_axial):
                if HeatFlux[i][j]>100:
                    heatflux_content[i] += str('%1.3f' % float(HeatFlux[i][j]))+' '
                elif HeatFlux[i][j]>10:
                    heatflux_content[i] += str('%1.4f' % float(HeatFlux[i][j]))+' '
                else:
                    heatflux_content[i] += str('%1.5f' % float(HeatFlux[i][j]))+' '
                # every line has 9 values
                if (j+1)%9==0:
                    heatflux_content[i] += '\n'




# this function write the coke.i file
def WriteCokeFile():
    # number of interpolations
    Interpo=10*(N_reactor_axial)-9
    for i in range(0,N_reactor):
        # reinitialize coke thickness content
        cokethickness_content[i]=''
        # interpolate cokethickness to the number of points required by coke.i
        CokeProfile=np.interp(np.linspace(0, Coilsim_Axial[N_reactor_axial-1], Interpo),Coilsim_Axial,CokeThickness[i])
        for j in range(0,Interpo):
            cokethickness_content[i] += '0 '+str(float(CokeProfile[j]))+'\n'
        # write the coke.i file
        FileCoke = open(WorkDir+'\\Projects\\USC\\Huajin'+str(i)+'\\coke.i','w')
        CokePre = '500    0.1\n'+str(Interpo)+'\n'
        FileCoke.write(CokePre+cokethickness_content[i])
        FileCoke.close()




# this function performs reactor simulations and fit CIP per reactor to get right COP
def SimulateReactors():
    FlagConvArray=zeros(N_reactor)
    FlagConv=0
    # CIP values
    x_low=zeros(N_reactor)
    x_high=zeros(N_reactor)
    x_old=zeros(N_reactor)
    x_new=zeros(N_reactor)
    # COP values
    y_old=zeros(N_reactor)
    # counter
    CounterCIP=zeros(N_reactor)
    # start iteration
    iteration=0
    while ((iteration<MaxCIPIteration) and (FlagConv==0)):
        iteration+=1
        # write simulation.txt
        f_simu = open(WorkDir+'\\Projects\\simulation.txt', 'w')
        simu_content = str(N_reactor-int(np.sum(FlagConvArray)))+'\n'+'0\n'
        for i in range(0,N_reactor):
            if FlagConvArray[i]==0:
                simu_content += 'USC\Huajin'+str(i)+'\n'
        f_simu.write(simu_content)
        f_simu.close()
        # update exp.txt and start reactor simulation
        for i in range(0,N_reactor):
            if FlagConvArray[i]==0:
                CounterCIP[i]+=1
                # write the exp.txt file
                FileExp = open(WorkDir+'\\Projects\\USC\\Huajin'+str(i)+'\\exp.txt','w')
                ExpPre = '1 \n1 \n'
                ExpNex = str(DilutionSteam)+'\n'+str(CIT)+' '
                ExpNex += str(CIP[i])+'\n'
                FileExp.write(ExpPre+heatflux_content[i]+'\n'+flowrate_content[i]+'\n'+ExpNex)
                FileExp.close()
        # change directory and run COILSIM1D
        os.chdir(WorkDir)
        subprocess.call(['Coilsim.exe'])
        os.chdir(os.path.pardir)
        # read COP and compare with COP setpoint to get new CIP
        for i in range(0,N_reactor):
            FileCSV = open(WorkDir+'\\Projects\\USC\\Huajin'+str(i)+'\\general_info.csv','rb')
            data = list(csv.reader(FileCSV,delimiter=','))
            COP[i]=float(data[N_reactor_axial][7].strip())
            if (abs(COP[i]-COPset))<CIPTreshold:
                FlagConvArray[i]=1
            FileCSV.close()
        '''# output the COP values
        OutputPE='            COP of all reactors: '
        for i in range(0,N_reactor):
            OutputPE += str(COP[i]) + ' '
        print OutputPE'''
        # update new CIP
        for i in range(0,N_reactor):
            if FlagConvArray[i]==0:
                # for COP smaller than 1.2
                if COP[i]<1.12:
                    x_low[i]=CIP[i]
                    CIP[i]=CIP[i]+1
                    CounterCIP[i]=0
                else:
                    # at least two iterations are needed
                    if CounterCIP[i]==1:
                        if COP[i]<COPset:
                            x_low[i]=CIP[i]
                            x_old[i]=CIP[i]
                            y_old[i]=COP[i]
                            CIP[i]=CIP[i]*1.01
                        else:
                            x_high[i]=CIP[i]
                            x_old[i]=CIP[i]
                            y_old[i]=COP[i]
                            CIP[i]=CIP[i]*0.99
                    else:
                        # set upper or lower limits for CIP
                        if COP[i]<COPset:
                            x_low[i]=CIP[i]
                        else:
                            x_high[i]=CIP[i]
                        # calculate new CIP based on the COP
                        if x_old[i]==CIP[i]:
                            x_new[i]=CIP[i]+(x_old[i]-CIP[i])/0.0001*(COPset-COP[i])
                        else:
                            x_new[i]=CIP[i]+(x_old[i]-CIP[i])/(y_old[i]-COP[i])*(COPset-COP[i])
                        x_old[i]=CIP[i]
                        y_old[i]=COP[i]
                        # 
                        if (x_low[i]!=0 and x_high[i]!=0):
                            # in the case of new CIP exceeds the lower limit
                            if x_new[i]<x_low[i]:
                                # avoid repeating of CIP in two iterations when the calculated x_new is always smaller than the lower limit
                                if CIP[i]==x_low[i]*1.01:
                                    CIP[i]=(x_low[i]+x_high[i])/2.0
                                else:
                                    CIP[i]=x_low[i]*1.01
                            # in the case of new CIP exceeds the upper limit
                            elif x_new[i]>x_high[i]:
                                # avoid repeating of CIP in two iterations when the calculated x_new is always larger than the upper limit
                                if CIP[i]==x_high[i]*0.99:
                                    CIP[i]=(x_low[i]+x_high[i])/2.0
                                else:
                                    CIP[i]=x_high[i]*0.99
                            else:
                                CIP[i]=x_new[i]
                        else:
                            CIP[i]=x_new[i]
        # iteraion is completed
        if (np.sum(FlagConvArray)==N_reactor):
            FlagConv=1
        print '            CIP iteration: '+ str(iteration)
        # output the CIP values
        OutputPE='            CIP of all reactors: '
        for i in range(0,N_reactor):
            OutputPE += str(CIP[i]) + ' '
        print OutputPE




# this function summarizes the reactor results
def WriteResults():
    # retrieve data from the results folders
    for i in range(0,N_reactor):
        # read general_info.csv
        FileCSV_info = open(WorkDir+'\\Projects\\USC\\Huajin'+str(i)+'\\general_info.csv','rb')
        data = list(csv.reader(FileCSV_info,delimiter=','))
        for j in range(0,N_reactor_axial):
            HeatFlux[i][j]=float(data[j+1][2].strip())
            ProcessT[i][j]=float(data[j+1][3].strip())
            TMT[i][j]=float(data[j+1][5].strip())
            ProcessP[i][j]=float(data[j+1][7].strip())
            CokingRate[i][j]=float(data[j+1][14].strip())
        FileCSV_info.close()
        # read results_summary.txt
        FileTXT = open(WorkDir+'\\Projects\\USC\\Huajin'+str(i)+'\\results_summary.txt','r')
        for j,line in enumerate(FileTXT):
            if j==13:
                HeatFluxTotal[i]=float(line.split()[4])
        FileTXT.close()
        # read yields.csv
        FileCSV_yield = open(WorkDir+'\\Projects\\USC\\Huajin'+str(i)+'\\yields.csv','rb')
        data = list(csv.reader(FileCSV_yield,delimiter=','))
        C2H4[i]=float(data[3][2].strip())
        C3H6[i]=float(data[7][2].strip())
        PE[i]=C3H6[i]/C2H4[i]
    
    # summarization of data to excel file 
    wb = xlwt.Workbook()
    # add sheets
    sheet = wb.add_sheet("ExternalWallTemperatures", cell_overwrite_ok=True)
    sheet2 = wb.add_sheet("ProcesgasTemperatures", cell_overwrite_ok=True)
    sheet3 = wb.add_sheet("Yields", cell_overwrite_ok=True)
    sheet4 = wb.add_sheet("ProcesgasPressure", cell_overwrite_ok=True)
    sheet5 = wb.add_sheet("CokingRate", cell_overwrite_ok=True)
    sheet6 = wb.add_sheet("CokeThickness", cell_overwrite_ok=True)
    sheet7 = wb.add_sheet("Heat Flux", cell_overwrite_ok=True)
    sheet8 = wb.add_sheet("Statistics", cell_overwrite_ok=True)
    # add axial title
    sheet.write(0,0,'Axial position [m]')
    sheet2.write(0,0,'Axial position [m]')
    sheet4.write(0,0,'Axial position [m]')
    sheet5.write(0,0,'Axial position [m]')
    sheet6.write(0,0,'Axial position [m]')
    sheet7.write(0,0,'Axial position [m]')
    # add axial title
    sheet.write(N_reactor_axial+3,0,'maxTMT')
    sheet5.write(N_reactor_axial+3,0,'maxCokingRate')
    sheet7.write(N_reactor_axial+3,0,'maxHeatFlux')
    # write profiles of TMT, Tgas, Pgas, coking rate, coke thickness and heat flux
    for i in range(0,N_reactor_axial):
        sheet.write(i+1,0,float(Coilsim_Axial[i]))
        sheet2.write(i+1,0,float(Coilsim_Axial[i]))
        sheet4.write(i+1,0,float(Coilsim_Axial[i]))
        sheet5.write(i+1,0,float(Coilsim_Axial[i]))
        sheet6.write(i+1,0,float(Coilsim_Axial[i]))
        sheet7.write(i+1,0,float(Coilsim_Axial[i]))
        for j in range(0,N_reactor):
            sheet.write(i+1,j+1,float(TMT[j][i]))
            sheet2.write(i+1,j+1,float(ProcessT[j][i]))
            sheet4.write(i+1,j+1,float(ProcessP[j][i]))
            sheet5.write(i+1,j+1,float(CokingRate[j][i]))
            sheet6.write(i+1,j+1,float(CokeThickness[j][i]))
            sheet7.write(i+1,j+1,float(HeatFlux[j][i]))
    # write maximum values and reactor coils numberss
    for j in range(0,N_reactor):
        sheet.write(0,j+1,'Reactor nr '+str(j+1))
        sheet2.write(0,j+1,'Reactor nr '+str(j+1))
        sheet4.write(0,j+1,'Reactor nr '+str(j+1))
        sheet5.write(0,j+1,'Reactor nr '+str(j+1))
        sheet6.write(0,j+1,'Reactor nr '+str(j+1))
        sheet7.write(0,j+1,'Reactor nr '+str(j+1))
        sheet.write(N_reactor_axial+3,j+1,float(max(TMT[j])))
        sheet5.write(N_reactor_axial+3,j+1,float(max(CokingRate[j])))
        sheet7.write(N_reactor_axial+3,j+1,float(max(HeatFlux[j])))
    # calculate the maximum CIP and TMT for run length simulation
    global MaxTMT
    global MaxCIP
    global MixingCupPE
    MaxTMT=np.max(TMT)
    MaxCIP=np.max(ProcessP[:][0])
    # write olefin yields and P/E
    sheet3.write(1,0,str("C2H4 [wt%]"))
    sheet3.write(2,0,str("C3H6 [wt%]"))
    sheet3.write(4,0,str("PE ratio [-]"))
    for j in range(0,N_reactor):
        sheet3.write(0,j+1,'Reactor nr '+str(j+1))
        sheet3.write(1,j+1,float(C2H4[j]))
        sheet3.write(2,j+1,float(C3H6[j]))
        sheet3.write(4,j+1,float(PE[j]))
    # write mass flow, mixing-cup P/E and tota absorbed heat
    sheet8.write(0,0,'Reactor')
    sheet8.write(1,0,'Mass flow [kg/h]')
    sheet8.write(2,0,'P/E')
    sheet8.write(3,0,'Total heat flux per reactor [kW]')
    sheet8.write(5,0,'Mixing cup average P/E')
    sheet8.write(6,0,'Total heat input to all reactors [kW]')
    PEWeighted=zeros(N_reactor)
    for j in range(N_reactor):
        sheet8.write(0,j+4,j+1)
        sheet8.write(1,j+4,FlowRate[j])
        sheet8.write(2,j+4,PE[j])
        sheet8.write(3,j+4,HeatFluxTotal[j])
        PEWeighted[j]=FlowRate[j]*PE[j]
    MixingCupPE=sum(PEWeighted)/sum(FlowRate)
    sheet8.write(5,4,float(MixingCupPE))
    TotalHeat=sum(HeatFluxTotal)
    sheet8.write(6,4,float(TotalHeat))
    # write other variables according to the simulation method
    if CoupledSim==1:
        sheet8.write(8,0,'Flue gas outlet temperature [K]')
        sheet8.write(9,0,'Heat absorbed (Furnace) [kW]')
        sheet8.write(11,0,'Fuel flow rate scaling ratio')
        sheet8.write(12,0,'Incident radiation scaling ratio')
        sheet8.write(8,4,float(T_fluegas))
        sheet8.write(9,4,float(Q_absorb))
        sheet8.write(11,4,float(FuelScalingRatio))
        sheet8.write(12,4,float(IncidentScalingRatio))
        wb.save(WorkDir+'\\Projects\\USC\\Reactor_Results_'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'_it'+str(IterationTMT)+'.xls')
    else:
        sheet8.write(8,0,'Heat flux scaling ratio')
        sheet8.write(8,4,float(HeatFluxScalingRatio))
        wb.save(WorkDir+'\\Projects\\USC\\Reactor_Results_'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'.xls')
    
    # make folder for results for this iteration (delete if already exists)
    if CoupledSim==1:
        Path = ResultDir+'\\'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'_it'+str(IterationTMT)
    else:
        Path = ResultDir+'\\'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)
    if  os.path.exists(Path): shutil.rmtree(Path)
    time.sleep(4.0)
    # copy files to the results folder
    shutil.copytree(WorkDir+'\\Projects\\USC',Path)
    # remove the summary excel file
    if CoupledSim==1:
        os.remove(WorkDir+'\\Projects\\USC\\Reactor_Results_'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'_it'+str(IterationTMT)+'.xls')
        if IterationTMT>1:
            os.remove(WorkDir+'\\Projects\\USC\\HuajinUSC_heatflux_'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'_it'+str(IterationTMT)+'.xls')
    else:
        os.remove(WorkDir+'\\Projects\\USC\\Reactor_Results_'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'.xls')
    print '            Results summary terminated successfully!'




# this function calculate the polynomial regression of the TMT
def GenerateWalltempCoefficient():
    # find out the division point of the inlet and outlet legs
    for i in range(0,N_reactor_axial):
        if Coilsim_Axial[i]>AxialDivPointLeg:
            mark=i
            break
    # assgin the corresponding furnace z-coordinates
    Furnace_Position=zeros(N_reactor_axial)
    for i in range(0,N_reactor_axial):
        if Coilsim_Axial[i] <= 9.15:
            Furnace_Position[i] = 11.609 - Coilsim_Axial[i]
        elif Coilsim_Axial[i] <= 10.16717:
            Furnace_Position[i] = 2.459 - 0.95*(Coilsim_Axial[i]-9.15)/1.01717
        elif Coilsim_Axial[i] <= 10.26717:
            Furnace_Position[i] = 11.67617 - Coilsim_Axial[i]
        elif Coilsim_Axial[i] <= AxialDivPointLeg:
            Furnace_Position[i] = 1.409 - 0.74*(math.sin((Coilsim_Axial[i]-10.26717)/0.74))
        elif Coilsim_Axial[i] <= 12.59195:
            Furnace_Position[i] = 0.74 - 0.74*math.cos((Coilsim_Axial[i]-11.42956)/0.74) + 0.669
        else:
            Furnace_Position[i] = Coilsim_Axial[i] - 12.59195 + 1.409
    # TMT values in Kelvin
    for i in range(0,N_reactor):
        TMTinKelvin=array(TMT[i])+273.15
        # regression
        TMTCoefInlet[i] = polyfit(Furnace_Position[:mark],TMTinKelvin[:mark],6)
        TMTCoefOutlet[i] = polyfit(Furnace_Position[mark:],TMTinKelvin[mark:],6)




# this function calculates the flue gas birdge wall temperature (K) and the total absorbed heat from furnace side(kW)
def FurnaceEstimation():
    global T_fluegas
    global Q_absorb
    # calculate the total absorbed heat by all reactor coils (kW/m2)
    Q_absorb=0.0
    HeatAbsorbedSingleCoil = zeros(N_reactor)
    for i in range(0,N_reactor):
        for j in range(0,2*N_furnace_points):
            if j==0:
                # the first segment
                HeatAbsorbedSingleCoil[i] += Heatflux_Furnace[i][j]*Furnace_Axial[j]*PI*InletID
            else:
                # inlet legs
                if Furnace_Axial[j-1]<AxialDivPointDiameter:
                    HeatAbsorbedSingleCoil[i] += Heatflux_Furnace[i][j-1]*(Furnace_Axial[j]-Furnace_Axial[j-1])*PI*InletID
                # outlet legs
                else:
                    HeatAbsorbedSingleCoil[i] += Heatflux_Furnace[i][j-1]*(Furnace_Axial[j]-Furnace_Axial[j-1])*PI*InletOD
        # the last segment
        HeatAbsorbedSingleCoil[i] += Heatflux_Furnace[i][2*N_furnace_points-1]*(AxialEndPoint-Furnace_Axial[2*N_furnace_points-1])*PI*InletOD
        '''print HeatAbsorbedSingleCoil[i]'''
        Q_absorb += HeatAbsorbedSingleCoil[i]
    Q_absorb=Q_absorb/1000
    
    # calculate furnace heat balance
    Q_release=Q_release_base*FuelScalingRatio
    Q_loss=LossRatio*Q_release
    Q_fluegas=Q_release-Q_loss-Q_absorb
    # flue gas molar flow rate (kmol/s)
    F_fluegas=F_fluegas_base*FuelScalingRatio/3600
    
    # parameters for enthalpy calculation
    N2_para_high=[2.95257637E+00,1.39690040E-03,-4.92631603E-07,7.86010195E-11,-4.60755204E-15,-9.23948688E+02,5.87188762E+00]
    N2_para_low=[3.53100528E+00,-1.23660988E-04,-5.02999433E-07,2.43530612E-09,-1.40881235E-12,-1.04697628E+03,2.96747038E+00]
    O2_para_high=[3.66096065E+00,6.56365811E-04,-1.41149627E-07,2.05797935E-11,-1.29913436E-15,-1.21597718E+03,3.41536279E+00]
    O2_para_low=[3.78245636E+00,-2.99673416E-03,9.84730201E-06,-9.68129509E-09,3.24372837E-12,-1.06394356E+03,3.65767573E+00]
    CO2_para_high=[4.63651110E+00,2.74145690E-03,-9.95897590E-07,1.60386660E-10,-9.16198570E-15,-4.90249040E+04,-1.93489550E+00]
    CO2_para_low=[2.35681300E+00,8.98412990E-03,-7.12206320E-06,2.45730080E-09,-1.42885480E-13,-4.83719710E+04,9.90090350E+00]
    H2O_para_high=[2.67703890E+00,2.97318160E-03,-7.73768890E-07,9.44335140E-11,-4.26899910E-15,-2.98858940E+04,6.88255000E+00]
    H2O_para_low=[4.19863520E+00,-2.03640170E-03,6.52034160E-06,-5.48792690E-09,1.77196800E-12,-3.02937260E+04,-8.49009010E-01]
    
    # iterative calculation of the flue gas bridge wall temperature
    b_upper=0.0
    b_lower=0.0
    LoopEnd=False
    while LoopEnd==False:
        E_N2=0.0
        E_O2=0.0
        E_CO2=0.0
        E_H2O=0.0
        # calculate enthalpy at current flue gas bridge wall temperature (T_fluegas)
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
        # calculate total enthalpy (kJ) of flue gas at bridge wall temperature (T_fluegas)
        # enthalpy difference (kJ/kmol) of flue gas between T_fluegas and reference temperature (298.15 K)
        E_N2=(E_N2*8.3145-0.0)*0.7193
        E_O2=(E_O2*8.3145-0.0)*0.0174
        E_CO2=(E_CO2*8.3145+393510.0)*0.0844
        E_H2O=(E_H2O*8.3145+241826.0)*0.1789
        E_total=E_N2+E_O2+E_CO2+E_H2O
        Q_fluegas_underT=F_fluegas*E_total
        delta_error=Q_fluegas_underT-Q_fluegas
        # converged
        if abs(delta_error)<0.1:
            LoopEnd=True
        # undate new flue gas bridge wall temperature (T_fluegas)
        else:
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
                print '!! warning: flue gas temperature T: '+ str(T_fluegas) + '  out of the range (200 - 6000 K)'




# calculate heat flux profile based on TMT and incident radiative heat flux
def GenerateHeatflux():
    heatflux_inlet = zeros((N_reactor,N_furnace_points))
    heatflux_outlet = zeros((N_reactor,N_furnace_points))
    for i in range(0,N_reactor):
        for j in range(0,N_furnace_points):
            # calculate TMT in Kelvin
            TempInlet=TMTCoefInlet[i][0]*x_position_inlet[j]**6.0+TMTCoefInlet[i][1]*x_position_inlet[j]**5.0+TMTCoefInlet[i][2]*x_position_inlet[j]**4.0+TMTCoefInlet[i][3]*x_position_inlet[j]**3.0+TMTCoefInlet[i][4]*x_position_inlet[j]**2.0+TMTCoefInlet[i][5]*x_position_inlet[j]+TMTCoefInlet[i][6]
            TempOutlet=TMTCoefOutlet[i][0]*x_position_outlet[j]**6.0+TMTCoefOutlet[i][1]*x_position_outlet[j]**5.0+TMTCoefOutlet[i][2]*x_position_outlet[j]**4.0+TMTCoefOutlet[i][3]*x_position_outlet[j]**3.0+TMTCoefOutlet[i][4]*x_position_outlet[j]**2.0+TMTCoefOutlet[i][5]*x_position_outlet[j]+TMTCoefOutlet[i][6]
            # calculate the external heat flux (kW/m^2)
            heatflux_inlet[i][j]=WallEmissivity*(IncidentR_inlet[i][j]*IncidentScalingRatio-StefanBoltzmann*TempInlet**4)*(1.0+ConvRadratio_inlet[i][j])
            heatflux_outlet[i][j]=WallEmissivity*(IncidentR_outlet[i][j]*IncidentScalingRatio-StefanBoltzmann*TempOutlet**4)*(1.0+ConvRadratio_outlet[i][j])
    # write excel file for internal heat flux profile
    writebook = xlwt.Workbook(encoding="utf-8")
    writesheet = writebook.add_sheet('heatflux_profile_innerwall')
    # convert external heat flux to internal heat flux (kW/m^2)
    for i in range(0,N_reactor):
        # reverse inlet heat flux values
        for j in range(0,N_furnace_points/2):
            TempVar=heatflux_inlet[i][j]
            heatflux_inlet[i][j]=heatflux_inlet[i][N_furnace_points-(j+1)]
            heatflux_inlet[i][N_furnace_points-(j+1)]=TempVar
        # write the title for each tube
        writesheet.write(0, i*3, 'Tube ' + str(i+1))
        writesheet.write(1, i*3, 'Axial Position(m)')
        writesheet.write(1, i*3+1, 'Heat Flux(W/m2)')
        for j in range(0,N_furnace_points):
            Heatflux_Furnace[i][j] = heatflux_inlet[i][j]*56.6/45.0
            if x_position_outlet[j] <= 1.409:
                Heatflux_Furnace[i][N_furnace_points+j] = heatflux_outlet[i][j]*56.6/45.0
            else:
                Heatflux_Furnace[i][N_furnace_points+j] = heatflux_outlet[i][j]*66.6/51.0
            # write the internal heatflux profile for each tube
            writesheet.write(j+2, i*3, Furnace_Axial[j])
            writesheet.write(j+2, i*3+1, Heatflux_Furnace[i][j])
            writesheet.write(N_furnace_points+(j+2), i*3, Furnace_Axial[N_furnace_points+j])
            writesheet.write(N_furnace_points+(j+2), i*3+1, Heatflux_Furnace[i][N_furnace_points+j])
    writebook.save(WorkDir+'\\Projects\\USC\\HuajinUSC_heatflux_'+CaseName+'_timestep'+str(TimeStep)+'_PEloop'+str(IterationPE)+'_it'+str(IterationTMT+1)+'.xls')




# this function read the heat flux (W/m^2) from the template folder and convert it to (kcal/m^2)
def GetHeatFlux():
    HeatFluxTemp=zeros(2*N_furnace_points)
    for i in range(0,N_reactor):
        for j in range(0,2*N_furnace_points):
            # convert heat flux to (kcal/m2)
            HeatFluxTemp[j] = Heatflux_Furnace[i][j]/1000.0/4.18400
    # interpolate heatflux to COILSIM axial positions
        HeatFlux[i]=np.interp(Coilsim_Axial, Furnace_Axial, HeatFluxTemp)












# ##############-------------------------- Main Function --------------------------############## #
#------------------------------ Read inputfile.txt ------------------------------#
variable=[None]*100
Input_file = open('inputfile.txt','r')
for i, line in enumerate(Input_file):
    if line!='\n':
        variable[i] = line.split()[0]
# assign values to the corresponding variables
# dir #
WorkDir=variable[1]            # work dir
ResultDir=variable[2]          # result sdir

# simulation option #
COILSIM_version=variable[6]     # coilsim version, v3.1
CaseName=variable[7]            # case name, original, coke, COT, PE
CoupledSim=int(variable[8])     # perform coupled simulation, 1.yes, 0.no
RunLengthSim=int(variable[9])   # perform runLength simulation, 1.yes, 0.no
ShootPE=int(variable[10])       # perform P/E shooting simulation, 1.yes, 0.no

# template #
TempDir=variable[14]             # template folder
FileReactor=variable[15]         # name of the reactor result template file in the folder
FileHeatFlux=variable[16]        # name of the heat flux template file
FileIncidentR=variable[17]       # name of the incident radiative heat flux template file

# base case condition #
T_fluegas_base=float(variable[21])  # flue gas birdge wall temperature (T_fluegas) in base case (K)
Q_release_base=float(variable[22])  # total heat release (Q_release) in base case (kW)
F_fluegas_base=float(variable[23])  # flue gas flow rate (F_fluegas) in base case (kmol/h)
FuelScalingRatio=float(variable[24])# fuel gas flow rate scaling factor

# run length simulation #
StartTimeStep=int(variable[28])     # initial time step (h)
TimeInterval=int(variable[29])      # time step interval of (h)
MaxTimeStep=int(variable[30])       # maximum run length time step
MaxTMTset=float(variable[31])       # end-of-run criteria TMT (C)
MaxCIPset=float(variable[32])       # end-of-run criteria CIP (atm)
CokeCorrelation=float(variable[33]) # coking rate scaling factor

# boundary condition #
DilutionSteam=float(variable[37])       # dilution steam
CIT=float(variable[38])                 # CIT (C)
COPset=float(variable[39])              # COP set value (atm)
MixingCupPEtarget=float(variable[40])   # mixing-up P/E set value (only for P/E shooting simulation)

# convergence #
TMTRelaxFactor=float(variable[44])      # TMT relaxation factor
IncidentRelaxFactor=float(variable[45]) # incident scaling relaxation factor
MaxPEIteration=int(variable[46])        # Maximum P/E iteration
MaxTMTIteration=int(variable[47])       # Maximum TMT iteration
MaxCIPIteration=int(variable[48])       # Maximum CIP iteration
PETreshold=float(variable[49])          # P/E convergence treshold
TMTTreshold=float(variable[50])         # TMT convergence treshold
CIPTreshold=float(variable[51])         # CIP convergence treshold
BalanceTreshold=float(variable[52])     # furnace heat balance treshold

# geometry info #
N_reactor=int(variable[56])          # number of the reactor coil
N_reactor_axial=int(variable[57])    # number of reactor axial points in COILSIM1D (two passes)
N_furnace_points=int(variable[58])   # number of reactor axial points in furnace (one pass)

# feedstock mass flow rate (kg/h)
variable[62]=variable[62].split(',')
FlowRate = zeros(N_reactor)
for i in range(0,N_reactor):
    FlowRate[i]=variable[62][i]
Input_file.close()

# other constants
PI=3.141592654
StefanBoltzmann=5.670367e-8     # Stefan-Boltzmann constant (W/m^2 T^4)
InletID=0.045                   # inner diameter of inlet leg (m)
InletOD=0.051                   # inner diameter of outlet leg (m)
AxialDivPointLeg=11.42956       # dividing point of the inlet and outlet legs (m)
AxialDivPointDiameter=12.59195  # dividing point of the inlet and outlet diameters (m)
AxialEndPoint=22.79195          # ending point of the reactor coil (m)
LossRatio=0.01                  # heat loss ratio through furance refractory
WallEmissivity=0.85             # reactor coil wall emissivity
CokeDensity=1600                # coke density (kg/m^3)

# value initialization #
T_fluegas=T_fluegas_base
Q_absorb=1.0
IncidentScalingRatio=1.0
HeatFluxScalingRatio=1.0
MaxTMT=0.0
MaxCIP=0.0
MixingCupPE=0.0
#------------------------------ Read inputfile.txt ------------------------------#




#------------------------------ Read template files ------------------------------#
# read incident radiative heat flux profiles (W/m^2)
x_position_inlet = zeros(N_furnace_points)
x_position_outlet = zeros(N_furnace_points)
IncidentR_inlet = zeros((N_reactor,N_furnace_points))
IncidentR_outlet = zeros((N_reactor,N_furnace_points))
ConvRadratio_inlet = zeros((N_reactor,N_furnace_points))
ConvRadratio_outlet = zeros((N_reactor,N_furnace_points))
GetInputVariables()


# read axial position, intitial coke thickness (m), CIP (atm), TMT (C)
Coilsim_Axial = zeros(N_reactor_axial)
CokeThickness = zeros((N_reactor,N_reactor_axial))
CokingRate = zeros((N_reactor,N_reactor_axial))
CIP = zeros(N_reactor)
COP = zeros(N_reactor)
TMT = zeros((N_reactor,N_reactor_axial))
TMTCoefInlet = zeros((N_reactor,7))
TMTCoefOutlet = zeros((N_reactor,7))
GetInitialValues()
# check if maxTMT and maxCIP are smaller than the stopping criterion
TimeStep=StartTimeStep
if RunLengthSim==1:
    if MaxTMT>MaxTMTset:
        print 'Program terminated:  initial TMT maximum: ' + str(MaxTMT) + ' C already exceeds the stopping criteria: ' + str(MaxTMTset) + ' C'
        exit()
    elif MaxCIP>MaxCIPset:
        print 'Program terminated:  initial CIP maximum: ' + str(MaxCIP) + ' atm already exceeds the stopping criteria: ' + str(MaxCIPset) + ' atm'
        exit()
    # assgin first time step
    TimeStep=StartTimeStep+TimeInterval


# read heat flux profiles (W/m^2) and convert it to (kcal/m2)
Furnace_Axial = zeros(2*N_furnace_points)
Heatflux_Furnace = zeros((N_reactor,2*N_furnace_points))
HeatFlux = zeros((N_reactor,N_reactor_axial))
GetInitialHeatFlux()
# the initial heat flux profile will be used for scaling up as the heat flux value will change during the calculation
''' using HeatFlux_ini=HeatFlux will make the two arrays relevant to each other when the element changes, i.e change at the same time '''
HeatFlux_ini = array(HeatFlux)
#------------------------------ Read template files ------------------------------#


#------------------------------ Initialization ------------------------------#
'''
# write simulation.txt
f_simu = open(WorkDir+'\\Projects\\simulation.txt', 'w')
simu_content = str(N_reactor)+'\n'+'0\n'
for i in range(0,N_reactor):
    simu_content += 'USC\Huajin'+str(i)+'\n'
f_simu.write(simu_content)
f_simu.close()
'''


# copy the files in the template dir to the work dir
if os.path.exists(WorkDir+'\\Projects\\USC'): shutil.rmtree(WorkDir+'\\Projects\\USC')
time.sleep(4.0)
shutil.copytree(TempDir,WorkDir+'\\Projects\USC')
# remove the summary excel file
os.remove(WorkDir+'\\Projects\\USC\\'+FileReactor)


# Prepare the content of flow rate, heat flux, and coke thickness to be written in exp.txt and coke.i
flowrate_content = [None]*N_reactor
heatflux_content = [None]*N_reactor
cokethickness_content = [None]*N_reactor
PrepareContent('FlowRate')
PrepareContent('HeatFlux')
WriteCokeFile()


# initialize the variables needed for writting the results
ProcessT = zeros((N_reactor,N_reactor_axial))
ProcessP = zeros((N_reactor,N_reactor_axial))
C2H4 = zeros(N_reactor)
C3H6 = zeros(N_reactor)
PE = zeros(N_reactor)
HeatFluxTotal = zeros(N_reactor)
#------------------------------ Initialization ------------------------------#








#------------------------------ Start simulation ------------------------------#
print ' ********************************************************************* '
print ' *********************** Coilsim version: ' + COILSIM_version + ' *********************** '
print ' ********************************************************************* '


# set variables for loop control
# run length simulation
if RunLengthSim==1:
    if CoupledSim==1:
        print '#### Start coupled run length simulation of case: (' + CaseName + ') ####'
    if CoupledSim==0:
        print '#### Start standalone run length simulation of case: (' + CaseName + ') ####'
    ExecTimeStepLoop=True
    ExecPEloop=True
    PEOnly=False
    OnceOnly=False
else:
    # P/E shooting simulation
    if ShootPE==1:
        if CoupledSim==1:
            print '#### Start coupled P/E shooting simulation of case: (' + CaseName + ') ####'
        if CoupledSim==0:
            print '#### Start standalone P/E shooting simulation of case: (' + CaseName + ') ####'
        ExecTimeStepLoop=False
        ExecPEloop=True
        PEOnly=True
        OnceOnly=False
    # standalone simulation
    else:
        if CoupledSim==1:
            print '#### Start coupled steady state simulation of case: (' + CaseName + ') ####'
        if CoupledSim==0:
            print '#### Start standalone steady state simulation of case: (' + CaseName + ') ####'
        ExecTimeStepLoop=False
        ExecPEloop=False
        PEOnly=False
        OnceOnly=True

# ------------------------------------------------ standalone simulation
if CoupledSim==0:
    # -------------------------------- start time step loop
    TimeStepLoopFin=False
    IterationTimeStep=1
    while TimeStepLoopFin==False:
        # print the simulation status
        if ExecTimeStepLoop==True:
            print ' -------- TMT criterion: ' + str(MaxTMTset) + ' -------- '
            print ' -------- CIP criterion: ' + str(MaxCIPset) + ' -------- '
            print 'Time step: ' + str(TimeStep)
        # calculate coke thickness and update the coke.i file for the current iteration
        if ExecTimeStepLoop==True:
            CokeThickness+=array(CokingRate)/(CokeDensity*1000)*TimeInterval*CokeCorrelation
            WriteCokeFile()
        
        # ---------------- start P/E loop
        PELoopConv=False
        IterationPE=1
        while PELoopConv==False:
            # print the simulation status
            if ExecPEloop==True:
                print ' ---- Mixing-cup P/E target: ' + str(MixingCupPEtarget) + ' ---- '
                print '    P/E iteration: ' + str(IterationPE)
            # perform reactor simulations (CIP loop) and write results
            SimulateReactors()
            WriteResults()
            # check the convergence of the P/E loop
            if (abs(MixingCupPE-MixingCupPEtarget)/MixingCupPEtarget)<PETreshold:
                PELoopConv=True
            if IterationPE==MaxPEIteration:
                PELoopConv=True
                print '!! warning: PE loop reaches maximum iteration times !!'
            # standalone steady state simulation has been completed
            if OnceOnly==True:
                TimeStepLoopFin=True
                PELoopConv=True
                print '#### Standalone steady state simulation of case: (' + CaseName + ') is completed ####'
            # adjust the heat flux scaling ratio according to the mixing-cup P/E
            if PELoopConv==False:
                PE_now=MixingCupPE
                if IterationPE==1:
                    HeatFluxScalingRatio_old=HeatFluxScalingRatio
                    HeatFluxScalingRatio+=0.01
                else:
                    HeatFluxScalingRatio_new=HeatFluxScalingRatio+(HeatFluxScalingRatio_old-HeatFluxScalingRatio)/(PE_old-PE_now)*(MixingCupPEtarget-PE_now)
                    HeatFluxScalingRatio_old=HeatFluxScalingRatio
                    HeatFluxScalingRatio=HeatFluxScalingRatio_new
                # assign new heat flux profile
                HeatFlux=array(HeatFlux_ini)*HeatFluxScalingRatio
                '''HeatFlux = [ [ j*HeatFluxScalingRatio for j in i ] for i in HeatFlux ]'''
                PrepareContent('HeatFlux')
                # update P/E and interation step
                PE_old=PE_now
                IterationPE+=1
                print '    Mixing-cup P/E: '+ str(MixingCupPE)
            # print the P/E results
            else:
                if ExecPEloop==True:
                    print '    P/E loop is converged (mixing-cup P/E: ' + str(MixingCupPE) + ', heat flux scaling ratio: ' + str(HeatFluxScalingRatio) + ')'
        # ---------------- end P/E loop
        
        # print out results of the current time step
        if ExecTimeStepLoop==True:
            print 'Time step: ' + str(TimeStep) + ' is converged (maximum TMT: ' + str(MaxTMT) + ', maximum CIP: ' + str(MaxCIP) + ')'
        # check the convergence of the time step loop
        if MaxTMT>MaxTMTset:
            TimeStepLoopFin=True
            Flag=1
        if MaxCIP>MaxCIPset:
            TimeStepLoopFin=True
            Flag=2
        if IterationTimeStep==MaxTimeStep:
            TimeStepLoopFin=True
            Flag=0
        # standalone P/E shooting simulation has been completed
        if PEOnly==True:
            TimeStepLoopFin=True
            print '#### Standalone P/E shooting simulation of case: (' + CaseName + ') is completed ####'
        # advance in one more time step
        if TimeStepLoopFin==False:
            TimeStep+=TimeInterval
            IterationTimeStep+=1
        # print the run length results
        else:
            if ExecTimeStepLoop==True:
                if Flag==1:
                    print '#### Standalone run length simulation of case: (' + CaseName + ') is completed due to TMT maximum ####'
                if Flag==2:
                    print '#### Standalone run length simulation of case: (' + CaseName + ') is completed due to CIP maximum ####'
                if Flag==0:
                    print '#### Standalone run length simulation of case: (' + CaseName + ') is completed due to time step maximum ####'
    # -------------------------------- end time step loop
# ------------------------------------------------ standalone simulation


# ------------------------------------------------ coupled simulation
if CoupledSim==1:
    # -------------------------------- start time step loop
    TimeStepLoopFin=False
    IterationTimeStep=1
    while TimeStepLoopFin==False:
        # print the simulation status
        if ExecTimeStepLoop==True:
            print ' -------- TMT criterion: ' + str(MaxTMTset) + ' -------- '
            print ' -------- CIP criterion: ' + str(MaxCIPset) + ' -------- '
            print 'Time step: ' + str(TimeStep)
        # calculate coke thickness and update the coke.i file for the current iteration
        if ExecTimeStepLoop==True:
            CokeThickness+=array(CokingRate)/(CokeDensity*1000)*TimeInterval*CokeCorrelation
            WriteCokeFile()
        
        # ---------------- start P/E loop
        PELoopConv=False
        IterationPE=1
        while PELoopConv==False:
            # print the simulation status
            if ExecPEloop==True:
                print ' ---- Mixing-cup P/E target: ' + str(MixingCupPEtarget) + ' ---- '
                print '    P/E iteration: ' + str(IterationPE)
            
            # -------- start TMT loop
            TMTLoopConv=False
            IterationTMT=1
            while TMTLoopConv==False:
                print '        TMT iteration: ' + str(IterationTMT)
                # perform reactor simulations (TMT loop) and generate wall temperature coefficient
                SimulateReactors()
                FurnaceEstimation()
                WriteResults()
                # at least two iterations are required to update the new TMT profile
                if IterationTMT>1:
                    # check the convergence of the TMT loop
                    ErrorTMT=abs(array(TMT)-array(TMT_old))
                    MaxErrorTMT=np.max(ErrorTMT)
                    print '        Maximum TMT error: '+ str(MaxErrorTMT)
                    if MaxErrorTMT<TMTTreshold:
                        TMTLoopConv=True
                    if IterationTMT==MaxTMTIteration:
                        TMTLoopConv=True
                        print '!! warning: TMT loop reaches maximum iteration times !!'
                    # adjust the incident radiative heat flux scaling ratio via correlation between IRHF and flue gas birdge wall temperature
                    if TMTLoopConv==False:
                        TMT=array(TMT)*TMTRelaxFactor+array(TMT_old)*(1.0-TMTRelaxFactor)
                # furnace heat balance calculation to update the heat flux profile
                if TMTLoopConv==False:
                    GenerateWalltempCoefficient()
                    
                    # ---- start furnace estimation loop
                    FurnaceHeatBalance=False
                    IterationFurnace=1
                    while FurnaceHeatBalance==False:
                        # start furnace estimation
                        FurnaceEstimation()
                        # at least two iterations are required to update the new flue gas birdge wall temperature
                        if IterationFurnace>1:
                            # check the convergence of the furnace estimation loop
                            BalanceError=abs(T_fluegas-T_fluegas_old)
                            if BalanceError<BalanceTreshold:
                                FurnaceHeatBalance=True
                        # update incident radiative heat flux and heat flux for the new iteration
                        if FurnaceHeatBalance==False:
                            # incident radiative heat flux profile
                            deltaT=T_fluegas-T_fluegas_base
                            IncidentScalingRatio_new=1.0+4*deltaT/T_fluegas_base+6*(deltaT/T_fluegas_base)**2+4*(deltaT/T_fluegas_base)**3+(deltaT/T_fluegas_base)**4
                            IncidentScalingRatio=IncidentScalingRatio_new*IncidentRelaxFactor+IncidentScalingRatio*(1-IncidentRelaxFactor)
                            # update furnace heat flux profile, flue gas bridge wall temperature and iteration step
                            GenerateHeatflux()
                            T_fluegas_old=T_fluegas
                            IterationFurnace+=1
                    # ---- end furnace estimation loop
                    
                    # update reactor heat flux profile and iteration step
                    GetHeatFlux()
                    PrepareContent('HeatFlux')
                    TMT_old=array(TMT)
                    IterationTMT+=1
                #
                else:
                    print '        TMT loop is converged (Flue gas bridge wall temperature (K): '+ str(T_fluegas) + ', incident radiative heat flux scaling ratio: ' + str(IncidentScalingRatio) + ')'
            # -------- end TMT loop
            
            # check the convergence of the P/E loop
            if (abs(MixingCupPE-MixingCupPEtarget)/MixingCupPEtarget)<PETreshold:
                PELoopConv=True
            if IterationPE==MaxPEIteration:
                PELoopConv=True
                print '!! warning: PE loop reaches maximum iteration times !!'
            # standalone steady state simulation has been completed
            if OnceOnly==True:
                TimeStepLoopFin=True
                PELoopConv=True
                print '#### Coupled steady state simulation of case: (' + CaseName + ') is completed ####'
            # adjust the heat flux scaling ratio according to the mixing-cup P/E
            if PELoopConv==False:
                PE_now=MixingCupPE
                if IterationPE==1:
                    FuelScalingRatio_old=FuelScalingRatio
                    FuelScalingRatio+=0.01
                else:
                    FuelScalingRatio_new=FuelScalingRatio+(FuelScalingRatio_old-FuelScalingRatio)/(PE_old-PE_now)*(MixingCupPEtarget-PE_now)
                    FuelScalingRatio_old=FuelScalingRatio
                    FuelScalingRatio=FuelScalingRatio_new
                # update P/E and interation step
                PE_old=PE_now
                IterationPE+=1
                print '    Mixing-cup P/E: '+ str(MixingCupPE)
            # print the P/E results
            else:
                if ExecPEloop==True:
                    print '    P/E loop is converged (mixing-cup P/E: ' + str(MixingCupPE) + ', fuel scaling ratio: ' + str(FuelScalingRatio) + ')'
        # ---------------- end P/E loop
        
        # print out results of the current time step
        if ExecTimeStepLoop==True:
            print 'Time step: ' + str(TimeStep) + ' is converged (maximum TMT: ' + str(MaxTMT) + ', maximum CIP: ' + str(MaxCIP) + ')'
        # check the convergence of the time step loop
        if MaxTMT>MaxTMTset:
            TimeStepLoopFin=True
            Flag=1
        if MaxCIP>MaxCIPset:
            TimeStepLoopFin=True
            Flag=2
        if IterationTimeStep==MaxTimeStep:
            TimeStepLoopFin=True
            Flag=0
        # standalone P/E shooting simulation has been completed
        if PEOnly==True:
            TimeStepLoopFin=True
            print '#### Coupled P/E shooting simulation of case: (' + CaseName + ') is completed ####'
        # advance in one more time step
        if TimeStepLoopFin==False:
            TimeStep+=TimeInterval
            IterationTimeStep+=1
        # print the run length results
        else:
            if ExecTimeStepLoop==True:
                if Flag==1:
                    print '#### Coupled run length simulation of case: (' + CaseName + ') is completed due to TMT maximum ####'
                if Flag==2:
                    print '#### Coupled run length simulation of case: (' + CaseName + ') is completed due to CIP maximum ####'
                if Flag==0:
                    print '#### Coupled run length simulation of case: (' + CaseName + ') is completed due to time step maximum ####'
    # -------------------------------- end time step loop