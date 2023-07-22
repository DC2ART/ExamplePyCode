# APX555 Audio precision
# BL
# 31.05.2022
# Source: API browser for APX500 v7.0
# It has Default Values of ini_file
# 03.08.2022: set_XLR_Unbal_Zout, set_BNC_Unbal_Zout

import clr
import logging
import pathlib
logger = logging.getLogger(__name__)

current_path = str(pathlib.Path(__file__).parent.resolve())

clr.AddReference("System.Drawing")
clr.AddReference("System.Windows.Forms")
# Add a reference to the APx AP
clr.AddReference(current_path + r"\audio_precision\dll\AudioPrecision.API2.dll")
clr.AddReference(current_path + r"\audio_precision\dll\AudioPrecision.API.dll")

import AudioPrecision.API as API


class APx555():
    def __init__(self, prj_File):
        self.prj_File = prj_File
        # 22.07.2022: To start the applicatin, set Bench Mode, Visibility etc. is not necessary: it is part of the prj_File (consistency of prj_File is checked by calling function):
        self.APx =API.APx500_Application(API.APxOperatingMode.BenchMode)
        #Open a project file with init settings (general and/or module-specific)
        self.OpenProject(self.prj_File) #disable only Debugging
        self.APx.Visible = True #For production: False, Background-process, visible in task-manager
        self.APx.Minimize()
        self.Set_Default_Config()
        #self.APx.timeout = 10.0 # in sec
       
        # Delete current Monitors/Meters  
        # print("Delete current Monitors/Meters") log
        MeterCount = self.APx.BenchMode.Meters.Count
        for meterIndex in range(MeterCount - 1, -1, -1):
            self.APx.BenchMode.Meters.Remove((meterIndex))

        #Add desired Monitors/Meters
        self.APx.BenchMode.Meters.Add(API.BenchModeMeterType.ThdNRatioMeter)
        self.APx.BenchMode.Meters.Add(API.BenchModeMeterType.RmsLevelMeter)
        self.APx.BenchMode.Meters.Add(API.BenchModeMeterType.FrequencyMeter)
        self.APx.BenchMode.Meters.Add(API.BenchModeMeterType.PeakLevelMeter)

        #Configure settling for desired Monitors/Meters
        print("Configure settling for desired Monitors/Meters")
        self.APx.BenchMode.Analyzer.Settling.Enabled = True
        self.APx.BenchMode.Analyzer.Settling.ReadingTimeout = 5
        self.APx.BenchMode.Analyzer.Settling.TrackFirst = True
        logging.basicConfig(filename="Error.log", format="%(asctime)s - %(message)s", level= logging.ERROR)
        
    #08.07.2022:
    def Get_SettledMeter_Reading(
        self, 
        Detector:API.SettlingMeterType = API.SettlingMeterType.PeakLevel, 
        Unit = "dBu"):
        #Get Settled readings from meters all at the same time.
        #Create an array of settling meter types
        
        #Number of settled meters has no effect on measurement duration
        SettledMeterTypes = (
           API.SettlingMeterType.ThdNRatio,   
           API.SettlingMeterType.ThdNLevel,
           API.SettlingMeterType.RmsLevel, 
           API.SettlingMeterType.Frequency,
           API.SettlingMeterType.PeakLevel
        )   

        #Pass an array of meter types to get the ISettledResultCollection 
        try:
            MeterReadings = self.APx.BenchMode.GetSettledMeterReadings(SettledMeterTypes)
            if not MeterReadings:
                raise cant_get_settled_meter_readings("Can't get settled meter readings")
            Values = MeterReadings[Detector].GetValues(Unit)
            return Values   
        except API.APException as e:
            print(__name__ + ": APException \n")
            logger.exception(e) # Loggs the exception
            raise
        
     #22.07.2022
    def OpenProject(self, prj_File):
        '''Opens a project file and starts the APx application. The path is absolute.\n
        The calling function checks for file-existance, consistency etc.\n
        OpenProject returns nothing (void)'''
        self.APx.OpenProject(prj_File)

    def get_valid_status(self, Bmetertype:API.BenchModeMeterType = API.BenchModeMeterType.FrequencyMeter):
        return self.APx.BenchMode.Meters.IsValid(Bmetertype)

    def get_LowerLimitStatus(self, BmeterType:API.BenchModeMeterType = API.BenchModeMeterType.FrequencyMeter):
        # Inputs: BenchModeMeterType, InputChannelIndex:
        # Error: No method matches given arguments for GetLimitSettings: (<class 'int'>, <class 'int'>)
        # Input only metertype: Error: The index is out of range of the collection. (index = 8, min=0, max=3)
        # Documentation: Input is BenchModeMeterType and not an Index
        return self.APx.BenchMode.Meters.GetLimitSettings(BmeterType).Lower.PassedLimitCheck

    def get_UpperLimitStatus(self, BmeterType:API.BenchModeMeterType = API.BenchModeMeterType.FrequencyMeter):
        return self.APx.BenchMode.Meters.GetLimitSettings(BmeterType).Upper.PassedLimitCheck

    def get_InpUnitList(self):
        #show and return List of available Units of Generator
        arr_units = self.APx.BenchMode.AnalogInputRanges.UnitList
        #output:Vrms,dBV,dBu,dBrA,dBrB,dBSPL1,dBSPL2,dBm,W
        print("APx-AnalogInputRanges: Available Units:")
        for unit in arr_units:
            print(str(unit).replace("\n", ", "))
        return(arr_units)

    def Set_Default_Config(self):
        '''Used for reset configuration'''
        self.MEAS_DELAY_SEC = 0.2 #Reference to ini-File: MEAS_DELAY_SEC = 0.2 
        self.MEAS_DURATION_SEC = 0.5 #Reference to ini-File: MEAS_DURATION_SEC = 0.5 
        self.APx_GEN_LEVEL_dBu = 0.00 #Reference to ini-File:APx_GEN_LEVEL_dBu = 0.00
        self.set_Meas_Delay(self.MEAS_DELAY_SEC )
        self.set_meas_Duration(self.MEAS_DURATION_SEC)
        self.OUTPUTGEN_OFF() # APx_GEN_ON = False
        self.set_XLR_Unbal_Zout(API.AnalogBalancedSourceImpedance.SourceImpedance_40) # APx_GENCONFIG = "UNBAL" + XLR
        self.set_GenFreq(5) # APx_GEN_FREQ_HZ = 5 HZ (0 Hz not possible)
        self.set_GenLevel(API.InputChannelIndex.Ch1, self.APx_GEN_LEVEL_dBu, "dBu") # APx_GEN_LEVEL_dBu = 0 for Ch1
        self.set_Zin(API.InputConnectorType.AnalogUnbalanced, API.InputChannelIndex.Ch1, API.AnalogInputTermination.InputTermination_Unbal_100k) # APx_INCONFIG = "UNBAL" , 100K, Ch1
        self.set_Highpass_Filter(API.HighpassFilterMode.DC, 0) # APx_FILTER =0  No High_Pass Filter, 0 HZ
        self.set_Lowpass_Filter(API.LowpassFilterModeAnalog.AdcPassband, 0) # No Lowpass_Filter, 0HZ
        self.set_AutoRange(True) # Autorange
        self.set_Weighting_Filter(API.SignalPathWeightingFilterType.wt_None)

    SH_Meters = [
        API.SettlingMeterType.ThdNRatio,
        API.SettlingMeterType.ThdNLevel,
        API.SettlingMeterType.RmsLevel,
        API.SettlingMeterType.PeakLevel,
        API.SettlingMeterType.Frequency
    ]

    # set Meas_Delay
    #08.07.2022
    def set_Meas_Delay(self, MEAS_DELAY_SEC):
        for meter in self.SH_Meters:
            self.APx.BenchMode.Analyzer.Settling[meter][API.InputChannelIndex.Ch1].DelayTime = MEAS_DELAY_SEC           
    
    # set Meas_Duration
    #08.07.2022
    def set_meas_Duration(self, MEAS_DURATION_SEC):
        for meter in self.SH_Meters:
            self.APx.BenchMode.Analyzer.Settling[meter][API.InputChannelIndex.Ch1].SettlingTime = MEAS_DURATION_SEC
        
    #Tolerance in %
    def set_Settling_Tolerance(self, TOL:float = 0.1):
        for meter in self.SH_Meters:
            self.APx.BenchMode.Analyzer.Settling[meter][API.InputChannelIndex.Ch1].Tolerance = TOL

    # LowpassfilterMode AdcPassband, Butterworth, Elliptic
    def set_Lowpass_Filter(self, Mode:API.LowpassFilterModeAnalog = API.LowpassFilterModeAnalog.AdcPassband, Freq:float = 0):
        self.APx.BenchMode.Setup.LowpassFilterAnalog = Mode
        self.APx.BenchMode.Setup.LowpassFilterFrequencyAnalog = Freq

    # HighpassfilterMode DC, AC, Butterword, Elliptic
    def set_Highpass_Filter(self, Mode:API.HighpassFilterMode = API.HighpassFilterMode.DC, Freq:float = 0):
        self.APx.BenchMode.Setup.HighpassFilter = Mode
        self.APx.BenchMode.Setup.HighpassFilterFrequency = Freq

    # FilterType  None, wt_A,...
    def set_Weighting_Filter(self, Type:API.SignalPathWeightingFilterType = API.SignalPathWeightingFilterType.wt_None):
        self.APx.BenchMode.Setup.WeightingFilter = Type

    # set Output_Configuration: OutputConnetorType AnalogUnbalances, ZoutValue(20ohm, 50ohm, 75ohm, 100ohm, 600om)
    def set_XLR_Unbal_Zout(self, Value:API.AnalogBalancedSourceImpedance = API.AnalogBalancedSourceImpedance.SourceImpedance_40):
        '''
        Set unbalanced output via *balanced settings* + underpoint *SingleEnded* to enable unbalanced via XLR-Cable. 
        If in the command OutputConnector.Type was set to "unbalanced", a BNC cable would have to be used.
        Note that impedance values are from balanced output i.e. 40R, 100R ...
        '''
        self.APx.BenchMode.Setup.OutputConnector.Type = API.OutputConnectorType.AnalogBalanced
        self.APx.BenchMode.Setup.AnalogOutput.CommonModeConfiguration = API.AnalogBalancedOutputConfigurationType.SingleEnded
        self.APx.BenchMode.Setup.AnalogOutput.BalancedSourceImpedance = Value

    # set Output_Configuration: OutputConnetorType AnalogUnbalances, ZoutValue(20ohm, 50ohm, 75ohm, 100ohm, 600om)
    def set_BNC_Unbal_Zout(self, Value:API.AnalogUnbalancedSourceImpedance = API.AnalogUnbalancedSourceImpedance.SourceImpedance_20):
        '''
        Set output impedance on BNC connector, note that impedance values of 20R, 50R etc. i.e. half of balanced output are possible, but BNC to XLR cable has to be used.
        '''
        self.APx.BenchMode.Setup.OutputConnector.Type = API.OutputConnectorType.AnalogUnbalanced
        self.APx.BenchMode.Setup.AnalogOutput.UnbalancedSourceImpedance = Value

    # set Output_Configuration: OutputConnetorType AnalogBalances, ZoutValue(40ohm, 100ohm, 150ohm, 200ohm, 600ohm)
    def set_Bal_Zout(self, Value:API.AnalogBalancedSourceImpedance = API.AnalogBalancedSourceImpedance.SourceImpedance_40):
        self.APx.BenchMode.Setup.OutputConnector.Type = API.OutputConnectorType.AnalogBalanced
        self.APx.BenchMode.Setup.AnalogOutput.BalancedSourceImpedance = Value

    def set_CMR_Zout(self, Impedance:API.AnalogBalancedSourceImpedance=API.AnalogBalancedSourceImpedance.SourceImpedance_40):
        self.set_Bal_Zout(Impedance)
        self.APx.BenchMode.Setup.AnalogOutput.CommonModeConfiguration = API.AnalogBalancedOutputConfigurationType.CMTST                       

    # set Input_Configuration: Channel(1, 2), Type(AnalogUnbalancesd, AnalogBalanced,..), Zinvalue(300ohm, 600ohm, Unbal_100k, Bal_200k)
    def set_Zin(
        self, 
        Type:API.InputConnectorType = API.InputConnectorType.AnalogBalanced, 
        Ch:API.InputChannelIndex = API.InputChannelIndex.Ch1, 
        ZinVal:API.AnalogInputTermination = API.AnalogInputTermination.InputTermination_300 ):
        self.APx.BenchMode.Setup.InputConnector.Type = Type
        self.APx.BenchMode.Setup.AnalogInput.SetTermination(Ch, ZinVal)

    def set_GenTrackFirstCh(self, bTrackState):
        self.APx.BenchMode.Generator.Levels.TrackFirstChannel = bTrackState

    # set Generator_Level:Ch1, Ch2, Reference to ini-File: APx_GEN_LEVEL_dBu
    def set_GenLevel(self, Ch:API.InputChannelIndex = API.InputChannelIndex.Ch1, Level:float = 0.0, sUnit:str = "V"):
        self.APx.BenchMode.Generator.Levels.Unit = sUnit
        try:
            self.APx.BenchMode.Generator.Levels.SetValue(Ch, Level)
        except:
            print("Error: set_GenLevel(self, Ch, Level): correct input Ch as int and Level as float and Unit as string\n Available Units: \n")
            self.get_GenUnitList()

    #18.07: Edit set_LowerLimit and set_UperLimit
    def set_LowerLimit(
         self, 
         meter:API.BenchModeMeterType = API.BenchModeMeterType.PeakLevelMeter, 
         Input:API.APxInputSelection = API.APxInputSelection.Input1,
         Ch:API.InputChannelIndex = API.InputChannelIndex.Ch1, 
         Limit:float = 0.00,
         Unit:str = "V"):
         self.APx.BenchMode.Meters.GetLimitSettings(meter, Input).Lower.SetValue(Ch, Limit, Unit)

    def set_UpperLimit(
         self, 
         meter:API.BenchModeMeterType = API.BenchModeMeterType.PeakLevelMeter, 
         Input:API.APxInputSelection = API.APxInputSelection.Input1,
         Ch:API.InputChannelIndex = API.InputChannelIndex.Ch1, 
         Limit:float = 0.00,
         Unit:str = "V"):
         self.APx.BenchMode.Meters.GetLimitSettings(meter, Input).Upper.SetValue(Ch, Limit, Unit)

    def get_GenUnitList(self):
        #show and return List of available Units of Generator
        arr_units = self.APx.BenchMode.Generator.Levels.UnitList
        #output:Vrms,Vp,Vpp,dBV,dBu,dBrG,dBm,W
        print("APx-Generator: Available Units:")
        for unit in arr_units:
            print(str(unit).replace("\n", ", "))
        return(arr_units)
    
    # set Generator_frequenz, Reference to ini-File: APx_GEN_FREQ_HZ
    def set_GenFreq(self, APx_GEN_FREQ_HZ):
        self.APx.BenchMode.Generator.Frequency.Value = float(APx_GEN_FREQ_HZ)

    # set Output of Generator ON
    def OUTPUTGEN_ON(self):
        self.APx.BenchMode.Generator.On = True


    # set Output of Generator ON        
    def OUTPUTGEN_OFF(self):
        self.APx.BenchMode.Generator.On = False

    # set InputRange in AutoRange  
    def set_AutoRange(self, Checkmark):
        self.APx.BenchMode.AnalogInputRanges.AutoRange = Checkmark
         
    # reset Min, Max Memory
    def reset_AutoScale(
        self,
        metertype:API.BenchModeMeterType = API.BenchModeMeterType.PeakLevelMeter, 
        channel:API.APxInputSelection = API.APxInputSelection.Input1):
        self.APx.BenchMode.Meters.ResetMinMax(metertype)
        self.APx.BenchMode.Meters.ResetMinMax(metertype, channel)

    # Code MaxMeter, SystemOne: RangeGain  
    def get_MaxMeterReading(
        self,
        metertype:API.BenchModeMeterType = API.BenchModeMeterType.PeakLevelMeter, 
        channel:API.APxInputSelection = API.APxInputSelection.Input1,
        sUnit:str = "Vrms"):
        dMaxReading = self.APx.BenchMode.Meters.GetMaxValues(metertype, channel, sUnit)
        for d in dMaxReading:
            print(f"MaxVal: {d}")
        return dMaxReading #Max Reading as Double[]

    def set_LowerLimit(self, meter, Input, Ch, Limit, Unit):
        self.APx.BenchMode.Meters.GetLimitSettings(meter, Input).Lower.SetValue(Ch, Limit, Unit)

#-----------------------------------------------------------------

# approjx_patch = current_path + r"\audio_precision\SH_APx_Template.approjx"
# print("APx File path: " + approjx_patch)
# APX1 = APx555(approjx_patch)
#22.07.2022, BL, Added and Checked CMR Setting, works: initially setup is Unbalanced output, setting changes correctly to balanced and then from 600Ohm to 40Ohm (as in TS23 of EOL Test product 1493)
# APX1.set_CMR_Zout(API.AnalogBalancedSourceImpedance.SourceImpedance_40)
# APX1.set_Unbal_Zout(API.AnalogUnbalancedSourceImpedance.SourceImpedance_20)
# APX1.set_Weighting_Filter(API.SignalPathWeightingFilterType.wt_A)
#22.07.2022, BL: unbalanced, balanced outputs: 
# set_Bal_Zout and set_Unbal_Zout: replaced bare int-Values by meaningful enums, tested and it works


# APX1.set_LowerLimit(API.BenchModeMeterType.PeakLevelMeter, 
#         API.APxInputSelection.Input1,
#         API.InputChannelIndex.Ch1, -81.00, "V")
# APX1.set_UpperLimit(API.BenchModeMeterType.PeakLevelMeter, 
#         API.APxInputSelection.Input1,
#         API.InputChannelIndex.Ch1, -84.00, "V")




#APX1 = APx555
# (r"C:\Users\Prozesstechnik\Documents\Programming\EOLTest\Testrun1\SH_APx_Template.approjx")
# APX1.set_XLR_Unbal_Zout()
'''
#------------------------------------------------------------
# APX1.set_Lowpass_Filter(API.LowpassFilterModeAnalog.Butterworth,80000) #LowpassfilterMode AdcPassband, Butterworth, Elliptic and 80kHZ
# APX1.set_Highpass_Filter(API.HighpassFilterMode.Elliptic, 20) #HighpassfilterMode DC, AC, Butterworth, Elliptic and 20HZ
# APX1.set_Weighting_Filter(API.SignalPathWeightingFilterType.wt_None) #FilterType None, A_wt,...
# APX1.(API.AnalogUnbalancedSourceImpedance.SourceImpedance_75) #20ohm, 50ohm, 75ohm, 100ohm, 600om
# APX1.set_Bal_Zout(API.AnalogBalancedSourceImpedance.SourceImpedance_600) #ZoutVal = 600ohm
# APX1.set_Zin(API.InputConnectorType.AnalogBalanced, API.InputChannelIndex.Ch2, API.AnalogInputTermination.InputTermination_Bal_200k) #Zinvalue 300ohm, 600ohm, Unbal_100k, Bal_200k
# APX1.set_GenFreq(13) # Frequenz = 13Hz
# APX1.set_Settling_Tolerance(1)

# 05.07.2022 (Demo Mode):
# Lowpass_Filter: set Elliptic and Frequency = 80KHZ   -->pass
# Highpass_Filter: set Butterword and Frequency = 20HZ -->pass
# Weighting_Filte: set A_wt                            -->pass
# Bal_Zout: set Balanced and Source Impedance = 600Ohm -->pass
# Zin: set Balanced and Value  = 200KOhm               -->pass
# GenLevel:  set for CH1  and Amplitude = -12dBu       -->pass
# GenFreq: set Value = 13HZ                            -->pass
# The following Measurment as running                  -->pass                            
# Delete current Monitors/Meters
# Configure settling for desired Monitors/Meters
# ThdNValues: -1.9760157391853366 -6.470031401290924
# RmsLevelValues: -4.736936911213648 2.0360859896205312
# PeakLevelValues: 0.7144478806294501 0.040595139376819134
#------------------------------------------------------------


# 04.07.2022 (Demo Mode):
# Delete current Monitors/Meters
# Configure settling for desired Monitors/Meters
# ThdNValues: -18.919756061847067 -19.73784103969973    
# RmsValues: 1.1152869097862217 1.488963929251656       
# PeakLevelValues: 0.7375497655011714 0.7007346856407821
        

#------------------------------------------------------------------------
# aus dem File RangeGain.py
APX1.set_GenLevel(API.InputChannelIndex.Ch1, 4.00, "Vrms")
APX1.set_GenLevel(API.InputChannelIndex.Ch2, 1.00, "Vrms")
APX1.get_MaxMeterReading(
    API.BenchModeMeterType.RmsLevelMeter, API.APxInputSelection.Input1, "Vrms"
    )
APX1.get_MaxMeterReading(
        API.BenchModeMeterType.PeakLevelMeter, API.APxInputSelection.Input1, "V"
        )


APX1.reset_AutoScale(
    API.BenchModeMeterType.RmsLevelMeter, API.APxInputSelection.Input1
    )
APX1.reset_AutoScale(
    API.BenchModeMeterType.PeakLevelMeter, API.APxInputSelection.Input1
    )
APX1.get_MaxMeterReading(
    API.BenchModeMeterType.RmsLevelMeter, API.APxInputSelection.Input1, "Vrms"
    )
APX1.get_MaxMeterReading(
        API.BenchModeMeterType.PeakLevelMeter, API.APxInputSelection.Input1, "V"
        )
                
#-------------------------------------------------------------------------
##08.07.2022: access channel 1 and 2 of returned values
def print_meas(meas, unit:str):
    i=1
    for val in meas:
        print(f"Value Channel{i}: " + str(val) + " " + unit)
        i+=1

#08.07.2022 start
APX1.OUTPUTGEN_ON()
tstart = time.time()
meas = APX1.Get_SettledMeter_Reading(API.SettlingMeterType.ThdNRatio, "%")
print_meas(meas, "%")

meas = APX1.Get_SettledMeter_Reading(API.SettlingMeterType.ThdNLevel, "dBu")
print_meas(meas, "dBu")

meas = APX1.Get_SettledMeter_Reading(API.SettlingMeterType.RmsLevel, "Vrms")
print_meas(meas, "Vrms")

meas = APX1.Get_SettledMeter_Reading(API.SettlingMeterType.Frequency, "Hz")
print_meas(meas, "Hz")


meas = APX1.Get_SettledMeter_Reading(API.SettlingMeterType.PeakLevel, "V")
print_meas(meas, "V")

tend = time.time()
print(f"Duration: {round(float(tend) - float(tstart),2)} s")  #6.63sec, 6.75sec

APX1.OUTPUTGEN_OFF()
#08.07.2022 end 
#--------------------------------------------------------------------------------                              

'''
'''
meas_ch1 = APX1.Get_SettledMeter_Reading(API.SettlingMeterType.RmsLevel, "dBu")

print(meas_ch1[0])
print(meas_ch1[1])
'''
