'''
SH Testrun v1.0, Product: 1493
linear flow: no sequencer, GUI
purpose: validate python modules of test-instruments + wiring harness + measurement results, identify necessary adjustments
# BL 14.07.2022
'''
# 29.07.2022: exec_test_step: Added step to open all relais, before closing the current ones in case the relais to be switched change
# 04.08.2022: TS00, exec_test_step: added offset current measurement

from logging import raiseExceptions
from time import sleep, ctime as tm_ctime
from os.path import getctime as os_getctime
from os.path import getmtime as os_getmtime
from sys import exit as sys_exit
import math

from Instruments.relay_assoc_1493 import RELAY_ASSOC_EOL as SH_REL
from Instruments.MeasTypes import MEAS_TYPE as SH_MEAS
# instances and init for instruments
from Instruments.APX_Ctrl_E1_6 import APx555
from Instruments.DAQ9x_E import DAQ903A, DAQ901A, DAQ970A
from Instruments.Keithley_2200_Ctrl_E import Keithley_2200

# Imports for APx555
import clr
clr.AddReference("System.Drawing") 
clr.AddReference("System.Windows.Forms")
# Add a reference to the APx API
clr.AddReference("C:\Program Files\Audio Precision\APx500 7.0\API\AudioPrecision.API2.dll") 
clr.AddReference("C:\Program Files\Audio Precision\APx500 7.0\API\AudioPrecision.API.dll")
import AudioPrecision.API as API

class REL_STATE():
    '''
    set self.RELAYS to defined states
    '''
    def set_NoChange(self):
        return "noChange"
    def set_NoRel(self):
        return "noRel"


class SH_1493_Test():
    '''EOL Test: HE-Art.Nr. 101493002 Customer-Art.Nr. "ME103C"
    Sennheiser Products EOL Test: Teststeps, init, reset, GUI etc. '''

    def __init__(self):
        self.init_PWR_SPLY() #connect to PWR1-3, set init-values, outputs on

        self.init_DAQ() #connect to DAQ and init Relay cards + Multimeter
        self.initAPx()
        self.REL_DELAY_SEC = 0.1 #Settling time after setting relais
        self.setREL = REL_STATE()
        self.MEAS_DELAY_SEC = 3.0 # for current offset, equals to TS01
        self.MEAS_DURATION_SEC = 0.5 # activate after debung
        self.pwr1_offset_mA = 0.00 # Current Offset for PWR1 powersupply
        # self.pwr1_offset_meas() #pwr1 voltage: same as in TS01, activate after debung

    def TS00(self):
        self.TS_DESC = "TS00\tChange Testobject"
        self.reset_TestStand()
        self.PWR1_VOUT = 56.0
        self.pwr1.set_Vout_Imax(self.PWR1_VOUT, self.IMAX1_A)
        print("TS000: Prüfling einlegen und Prüfadapter absenken") #notification: print and/or show in GUI

    def pwr1_offset_meas(self) -> float:
        '''
        Current-offset pwr1 @ noload = no DUT for offset-correction in current measurements pwr1
        '''
        input("D-Sub25 Stecker von Prüfadapter abziehen: Offset Strommessung") #TODO: transfer notification to GUI
        self.pwr1_offset_mA = self.pwr1.meas_curr(self.MEAS_DELAY_SEC, self.MEAS_DURATION_SEC)*1E3 #Offset is used in self.exec_test_step, when called in in TS001
        input("D-Sub25 Stecker von Prüfadapter aufstecken")        
        
        return self.pwr1_offset_mA


    def TS01(self):
        self.TS_DESC = "TS01\tStromaufnahme V1"
        self.RELAYS = self.setREL.set_NoChange()
        self.LimitMin = 3.90
        self.LimitMax = 4.20
        self.MeasUnit = "mA"
        self.MEAS_DELAY_SEC = 3.0
        self.MEAS_DURATION_SEC = 0.5
        self.MEAS_TYPE = SH_MEAS.get("AMP_PWR1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()
    
    def TS02(self):
        self.TS_DESC = "TS02\tDGND +56V; MP104"
        self.RELAYS = "REL_MP104"
        self.LimitMin = 11.8
        self.LimitMax = 12.1
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()

    def TS03(self):
        self.TS_DESC = "TS03\tDGND +46V; MP104"
        self.RELAYS = self.setREL.set_NoChange()
        self.PWR1_VOUT = 46.0
        self.pwr1.set_Vout_Imax(self.PWR1_VOUT, self.IMAX1_A)
        self.LimitMin = 11.8
        self.LimitMax = 12.1
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 3.0
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()

    def TS04(self):
        self.TS_DESC = "TS04\t+V1; MP108"
        self.RELAYS = "REL_MP108"
        self.LimitMin = 27.8
        self.LimitMax = 29.0
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()
    
    def TS05(self):
        self.TS_DESC = "TS05\t+V3; MP102"
        self.RELAYS = "REL_MP102"
        self.LimitMin = 11.2
        self.LimitMax = 11.5
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()
    
    def TS06(self):
        self.TS_DESC = "TS06\tDC_OUT; MP103"
        self.RELAYS = "REL_MP103"
        self.LimitMin = 4.9
        self.LimitMax = 5.6
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()

    def TS07(self):
        self.TS_DESC = "TS07\tXLR1; MP107"
        self.RELAYS = "REL_MP107"
        self.LimitMin = -0.010
        self.LimitMax = 0.010
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()

    def TS08(self):
        self.TS_DESC = "TS08\tXLR2 minus XLR3"
        self.RELAYS = "REL_MP106"
        self.LimitMin = -0.015
        self.LimitMax = 0.015
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 3.0
        self.MEAS_DURATION_SEC = 0.5        
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()

    def TS09(self):
        #set relais
        self.TS_DESC = "TS09\t-VCS zehntel; MP205/10"
        self.RELAYS = 'REL_VCS/10'
        self.LimitMin = -6.50
        self.LimitMax = -6.20
        self.MeasUnit = "V"
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_DURATION_SEC = 0.5
        self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1")
        #APx_GEN_ON = False #APx Gen is still off
        self.exec_test_step()

    def TS10(self):
        #set relais
        self.TS_DESC = "TS10\tClock; MP201"
        self.RELAYS = 'REL_Clock'
        self.LimitMin = 66000.00
        self.LimitMax = 76000.00
        self.MeasUnit = "Hz"
        self.MEAS_DURATION_SEC = 1
        self.MEAS_DELAY_SEC = 0.2
        self.MEAS_TYPE = SH_MEAS.get("APX_FUNC") #remains until testend constant
        #Settling Parameters
        self.APX.set_meas_Duration(self.MEAS_DURATION_SEC)
        self.APX.set_Meas_Delay(self.MEAS_DELAY_SEC)
        self.benchDetector = API.BenchModeMeterType.FrequencyMeter
        self.settlDetector = API.SettlingMeterType.Frequency
        # self.APxInput, self.APxChannel: #unchanged from init
        HI_PASS_FILTER_FREQU_HZ = 10
        LO_PASS_FILTER_FREQU_HZ = 5E5 #500kHz
        #FilterType is constant and is defined at init
        self.APX.set_Highpass_Filter(self.HiPassFilterType, HI_PASS_FILTER_FREQU_HZ)
        self.APX.set_Lowpass_Filter(self.LoPassFilterType, LO_PASS_FILTER_FREQU_HZ)
        # self.APX.OUTPUTGEN_ON() #Output is off
        self.exec_test_step()

    def TS11(self):
        self.TS_DESC = "TS11\tMW als VAR"
        self.RELAYS = "REL_GEN"
        self.LimitMin = -2.40
        self.LimitMax = -1.90
        self.MeasUnit = "dBu"
        self.MEAS_DURATION_SEC = 0.5
        self.MEAS_DELAY_SEC = 0.2
        # self.MEAS_TYPE = SH_MEAS.get("APX_FUNC") no change
        # APx Generator Parameters
        APx_GEN_LEVEL_dBu = 0.00
        APx_GEN_FREQ_HZ = 5000 #in Hz
        #APx Settling Parameters
        self.APX.set_meas_Duration(self.MEAS_DURATION_SEC)
        #APX MEAS_DELAY_SEC: no change to prior teststep, no setting needed
        self.benchDetector = API.BenchModeMeterType.RmsLevelMeter
        self.settlDetector = API.SettlingMeterType.RmsLevel
        # settings specific to this Teststep
        self.APX.set_GenLevel(self.APxChannel, APx_GEN_LEVEL_dBu, "dBu")
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)
        # start measurement & store measurement in variable GEN_VAR 
        self.APX.OUTPUTGEN_ON()
        self.GEN_VAR = self.exec_test_step()[0]

    def TS12(self):
        self.TS_DESC = "TS12\tVerst. 1k Last"
        self.RELAYS = "REL_GEN, REL_Last_1K"
        self.LimitMin = -20.60
        self.LimitMax = -20.20
        self.MeasUnit = "dBu"
        # Frequency 5kHz, Detector RMS, unchanged

        # activate after debug
        APx_GEN_LEVEL_dBu = -(20.00 + float(self.GEN_VAR))
        self.APX.set_GenLevel(self.APxChannel, APx_GEN_LEVEL_dBu, "dBu")

        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS13(self):
        self.TS_DESC = "TS13\tFrqG. 1k Last"
        self.RELAYS = "REL_GEN, REL_Last_1K"
        self.LimitMin = -21.50
        self.LimitMax = -21.00        
        self.MeasUnit = "dBu"
        #APx_GEN_LEVEL_dBu = -(20.00 + float(self.GEN_VAR)) #no change from last step
        self.APx_GEN_FREQ_HZ = 20
        self.APX.set_GenFreq(self.APx_GEN_FREQ_HZ)
        # Detector RMS, unchanged
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS14(self):
        self.TS_DESC = "TS14\tFrqG. 20Hz"
        self.RELAYS = "REL_GEN"
        self.LimitMin =  -20.20
        self.LimitMax =  -20.00
        self.MeasUnit = "dBu"
        #APx_GEN_LEVEL_dBu = -(20.00 + float(self.GEN_VAR)) #no change from last step
        # Frequency is unchanged
        # Detector RMS, unchanged
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()
    
    def TS15(self):
        self.TS_DESC = "TS15\tFrqG. 150kHz"
        self.RELAYS = self.setREL.set_NoChange()
        self.LimitMin =  -21.00
        self.LimitMax =  -20.40
        self.MeasUnit = "dBu"
        #APx_GEN_LEVEL_dBu = -(20.00 + float(self.GEN_VAR)) #no change from last step
        APx_GEN_FREQ_HZ = 1.5E5
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)
        # Detector RMS, unchanged
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS16(self):
        self.TS_DESC = "TS16\tFrqG. 150kHz C(10nF)"
        self.RELAYS = "REL_GEN, REL_Last_10n"
        self.LimitMin =  -21.00
        self.LimitMax =  -20.20
        self.MeasUnit = "dBu"
        #APx_GEN_FREQ_HZ # no change from last step
        #APx_GEN_LEVEL_dBu  # no change from last step
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS17(self):
        self.TS_DESC = "TS17\tFrqG. 150kHz C(47nF)"
        self.RELAYS = "REL_GEN, REL_Last_47n"
        self.LimitMin =  -27.00
        self.LimitMax =  -25.80
        self.MeasUnit = "dBu"
        #APx_GEN_FREQ_HZ # no change from last step
        #APx_GEN_LEVEL_dBu  # no change from last step
        # Detector RMS, unchanged
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()


    def TS18(self):
        self.TS_DESC = "TS18\tTHDN Ratio C=47n 5kHz"
        self.RELAYS = self.setREL.set_NoChange()
        self.LimitMin =  -70.00
        self.LimitMax =  -58.00
        self.MeasUnit = "dB"
        self.benchDetector = API.BenchModeMeterType.ThdNRatioMeter
        self.settlDetector = API.SettlingMeterType.ThdNRatio
        APx_GEN_FREQ_HZ = 5000
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)

        # deactivate for debug, activate after debug
        APx_GEN_LEVEL_dBu = 11.00 - float(self.GEN_VAR)
        self.APX.set_GenLevel(self.APxChannel, APx_GEN_LEVEL_dBu, "dBu") # for Generator, Unit: not self.MeasUnit, but "dBu"

        HI_PASS_FILTER_FREQU_HZ = 22
        LO_PASS_FILTER_FREQU_HZ = 80E3
        self.APX.set_Highpass_Filter(self.HiPassFilterType, HI_PASS_FILTER_FREQU_HZ)
        self.APX.set_Lowpass_Filter(self.LoPassFilterType, LO_PASS_FILTER_FREQU_HZ)
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS19(self):
        self.TS_DESC = "TS19\tTHDN Ratio RL=1k 5kHz"
        self.RELAYS = "REL_GEN, REL_Last_1K"
        self.LimitMin =  -75.00
        self.LimitMax =  -56.00
        self.MeasUnit = "dB"
        # Detectors, Frequency, Generator-Level unchanged
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS20(self):
        self.TS_DESC = "TS20\tTHDN Ratio  RL=1k 25Hz"
        self.RELAYS = self.setREL.set_NoChange()
        self.LimitMin =  -85.00
        self.LimitMax =  -68.00
        self.MeasUnit = "dB"
        self.benchDetector = API.BenchModeMeterType.ThdNRatioMeter
        self.settlDetector = API.SettlingMeterType.ThdNRatio
        APx_GEN_FREQ_HZ = 25
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)

        # deactivate for debug, activate after debug
        APx_GEN_LEVEL_dBu = 11.00 - float(self.GEN_VAR)
        self.APX.set_GenLevel(self.APxChannel, APx_GEN_LEVEL_dBu, "dBu")

        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS21(self):
        self.TS_DESC = "TS21\tTHDN Ratio  Anmodulation + V2"
        self.RELAYS = "REL_GEN, REL_Last_1K, REL_Anmod"
        self.LimitMin =  -40.00
        self.LimitMax =  -27.00
        self.MeasUnit = "dBu"
        self.benchDetector = API.BenchModeMeterType.RmsLevelMeter
        self.settlDetector = API.SettlingMeterType.RmsLevel
        APx_GEN_FREQ_HZ = 25
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)

        # activate after debug
        APx_GEN_LEVEL_dBu = 6.00
        self.APX.set_GenLevel(self.APxChannel, APx_GEN_LEVEL_dBu, "dBu")

        HI_PASS_FILTER_FREQU_HZ = 10
        LO_PASS_FILTER_FREQU_HZ = 22E3
        self.APX.set_Highpass_Filter(self.HiPassFilterType, HI_PASS_FILTER_FREQU_HZ)
        self.APX.set_Lowpass_Filter(self.LoPassFilterType, LO_PASS_FILTER_FREQU_HZ)
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()


    def TS22(self):
        self.TS_DESC = "TS22\tEingangsbeschaltung"
        self.RELAYS = "REL_Einsp"
        self.LimitMin =  -61.00
        self.LimitMax =  -53.00
        self.MeasUnit = "dBu"
        # Detectors unchanged
        #Frequency unchanged from last teststep
        #Generator Level unchanged from last teststep
        #Filter Hi-/Lo-Pass unchanged from last teststep
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()


    def TS23(self):
        self.TS_DESC = "TS23\tSymmetrie 5kHz"
        self.RELAYS = "REL_Sym"
        self.LimitMin =  -80.00
        self.LimitMax =  -50.00
        self.MeasUnit = "dBu"
        # Detectors RMS, unchanged
        #CMR configuration, 40R Impedance
        self.APX.set_CMR_Zout(API.AnalogBalancedSourceImpedance.SourceImpedance_40)
        APx_GEN_FREQ_HZ = 5E3
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)
        APx_GEN_LEVEL_dBu = 0.00
        self.APX.set_GenLevel(self.APxChannel, APx_GEN_LEVEL_dBu, "dBu")
        HI_PASS_FILTER_FREQU_HZ = 22
        LO_PASS_FILTER_FREQU_HZ = 22E3
        self.APX.set_Highpass_Filter(self.HiPassFilterType, HI_PASS_FILTER_FREQU_HZ)
        self.APX.set_Lowpass_Filter(self.LoPassFilterType, LO_PASS_FILTER_FREQU_HZ)
        self.APX.OUTPUTGEN_ON()
        self.exec_test_step()

    def TS24(self):
        self.TS_DESC = "TS24\tSymmetrie 100Hz"
        self.RELAYS = "noChange"
        self.LimitMin =  -80.00
        self.LimitMax =  -50.00
        self.MeasUnit = "dBu"
        # Detectors RMS, unchanged
        #CMR configuration, 40R Impedance unchanged from last test step
        #GEN_LEVEL unchanged from last test step
        #Hi- Lo- Pass Filter unchanged from last test step
        APx_GEN_FREQ_HZ = 100
        self.APX.set_GenFreq(APx_GEN_FREQ_HZ)
        self.APX.OUTPUTGEN_ON() #OUTPUTGEN_ON cannot be in exec_test_step because in TS25/TS26 OUTPUTGEN has to be OFF
        self.exec_test_step()


    def TS25(self):
        self.TS_DESC = "TS25\tRauschen; quasi Quasi-Peak"
        self.RELAYS = "REL_PV40A"
        self.LimitMin =  self.dBu_to_V(-84.00) #APx PeakLevel supports only V as Unit
        self.LimitMax =  self.dBu_to_V(-81.00)
        self.MeasUnit = "V" #Peak Level: in APx nur V als Unit, später in dBu umrechnen, entsprechend Prüfplan
        #Generator configuration CMR, Impedance, Level, Frequency is not relevant in this teststep and these settings are not changed
        self.APX.OUTPUTGEN_OFF()
        self.benchDetector = API.BenchModeMeterType.PeakLevelMeter
        self.settlDetector = API.SettlingMeterType.PeakLevel
        self.APX.set_Weighting_Filter(API.SignalPathWeightingFilterType.wt_A)
        HI_PASS_FILTER_FREQU_HZ = 22
        LO_PASS_FILTER_FREQU_HZ = 22E3
        self.APX.set_Highpass_Filter(self.HiPassFilterType, HI_PASS_FILTER_FREQU_HZ)
        self.APX.set_Lowpass_Filter(self.LoPassFilterType, LO_PASS_FILTER_FREQU_HZ)
        self.V_to_dBu(0.666) #deactivate for debug, Dummy Value, 
        # Activate for production: 
        self.V_to_dBu(self.exec_test_step()[0]) #Measurement result from APx: in V, Limits in dBu
        #TODO: Logging
        # self.LimitMin =  -84.00 #for logging in dBu, not in V
        # self.LimitMax =  -81.00
        # self.MeasUnit = "V"


    def TS26(self):
        self.TS_DESC = "TS26\tRest-HF"
        self.RELAYS = self.setREL.set_NoRel() # All Relais switched Off
        self.LimitMin =  -105.00
        self.LimitMax =  -95.00
        self.MeasUnit = "dBu"
        #Generator configuration CMR, Impedance, Level, Frequency is not relevant in this teststep and these settings are not changed
        self.APX.OUTPUTGEN_OFF()
        self.benchDetector = API.BenchModeMeterType.RmsLevelMeter
        self.settlDetector = API.SettlingMeterType.RmsLevel
        self.APX.set_Weighting_Filter(API.SignalPathWeightingFilterType.wt_None)
        HI_PASS_FILTER_FREQU_HZ = 400
        LO_PASS_FILTER_FREQU_HZ = 5E5
        self.APX.set_Highpass_Filter(self.HiPassFilterType, HI_PASS_FILTER_FREQU_HZ)
        self.APX.set_Lowpass_Filter(self.LoPassFilterType, LO_PASS_FILTER_FREQU_HZ)
        self.exec_test_step()

    def TS27(self):
        self.TS_DESC = "TS27\tTest End"
        print("="*70)


    def reset_TestStand(self):
        self.DAQ.REL.SPON_All()
        # Pwr1,2,3 stay on all the time
        self.APX.OUTPUTGEN_OFF() #leave all other APx settings untouched
        print("Reset Teststand performed")

    def load_APxTemplate(self):
        '''Check consistency of project File: presence, \"created\" and \"modified\" time stamp\n
        In case of error raise exception and exit program with notification.\n
        In case of no errors start APx application with project file:\n
        self.APX = APx555(APx_Proj_Template)
        '''
        APx_Proj_Template = r"C:\Users\Prozesstechnik\Documents\Programming\EOLTest\Testrun1\SH_APx_Template.approjx" # APx BenchMode Template with all relevant init settings
        #reference template-timestamps for creation and modification:
        TM_CREATD = 'Mon Jul 18 18:08:58 2022'
        TM_MODIFD = 'Tue May 17 13:00:43 2022'

        # get time-stamps and convert from EPOCH format to timestamp
        try:
            # get time-stamp "created", "modified"
            # in case FileNotFound: inbuilt exception is raised
            tm_creatd = tm_ctime(os_getctime(APx_Proj_Template)) 
            tm_modifd = tm_ctime(os_getmtime(APx_Proj_Template))

            try: 
                if (tm_creatd != TM_CREATD or tm_modifd != TM_MODIFD):
                    raise ValueError 
                else:
                    print(f"APx template in {APx_Proj_Template} has been loaded successfully! Timestamps were checked and are OK, the APx application is started") # print or log
                    self.APX = APx555(APx_Proj_Template) #performs APx init
            except ValueError:
                exit_msg = f"Error while loading APx template in {APx_Proj_Template}:\n\"created\" and \"modified\" timestamps of the template don't match\nPlease check how the Template was changed \nThe program is aborted\nIf the changes in the template are valid, please adjust references for \"created\" and \"modified\" in this program"
                sys_exit(exit_msg) #exit program
        except FileNotFoundError:
            exit_msg = f"APx template in {APx_Proj_Template}: Das System kann die angegebene Datei nicht finden"
            sys_exit(exit_msg) #exit program
        except:
            exit_msg = f"Fehler beim Laden des APx templates in {APx_Proj_Template}\n"
            sys_exit(exit_msg)


    # general tasks in most teststeps
    def exec_test_step(self):
        '''set relays, trigger measurement, check measurement result\n
        return measurment value and status'''

        # only set Relais if they haven't been set already
        # translate Relay-Name to its int-Value, None if Name not found
        if self.RELAYS == self.setREL.set_NoRel():
            self.DAQ.REL.SPON_All()
        elif self.RELAYS == self.setREL.set_NoChange():
            pass
        elif self.RELAYS != "" and self.RELAYS != None:
            # in case self.RELAYS holds more than one relay
            # e.g. self.RELAYS = "REL_GEN, REL_Last_1K"
            # SH_REL.get will return None because it can find only one relay
            # solution: split and iterate over list, this works for one item, too
            
            arr_rel = []
            for item in self.RELAYS.split(","):
                rel = SH_REL.get(item.replace(" ", ""))
                if rel != None:
                    arr_rel.append(str(rel))
                else:
                    print("exec_test_step: unexpected error at relay string: received None from SH_REL.get, expected int ") #print or log
    
            for i in range(0,len(arr_rel)):
                if i == 0:
                   self.DAQ.REL.Set_XClose(arr_rel[i]) #close first relay in list exclusively, i.e. open all other relays. If the arr_rel contains more relais: the next relays are added non-exclusively
                else:
                    self.DAQ.REL.Set_Close(arr_rel[i]) #close non-exclusive, i.e. close relay and leave all other relays as is
            sleep(self.REL_DELAY_SEC)
        else:
            print("exec_test_step: unknown Relais string")

        
        #initialize return parameters, in case measurement is skipped, return parameters are unchanged
        meas_ch1 = "init"
        step_result = "NoResult" 

        #APX measurement
        if self.MEAS_TYPE == SH_MEAS.get("AMP_PWR1"):
            meas_ch1 = (self.pwr1.meas_curr(self.MEAS_DELAY_SEC, self.MEAS_DURATION_SEC))*1000 - self.pwr1_offset_mA #result in mA with offset correction


        elif self.MEAS_TYPE == SH_MEAS.get("VOLT_DMM_CH1"):
            DMMchannel = "01"
            meas_ch1 = self.DAQ.DMM.Meas_Volt_DC(DMMchannel, self.MEAS_DURATION_SEC, self.MEAS_DELAY_SEC) #Result in V
            self.MeasUnit = "V"
            #Tested in Testmodule DAQ970.py
     
        elif self.MEAS_TYPE == SH_MEAS.get("APX_FUNC"):
            self.APX.set_LowerLimit(self.benchDetector, self.APxInput, self.APxChannel, self.LimitMin, self.MeasUnit)
            self.APX.set_UpperLimit(self.benchDetector, self.APxInput, self.APxChannel, self.LimitMax, self.MeasUnit)


            # start settled measurement
            meas_ch1 = self.APX.Get_SettledMeter_Reading(
                self.settlDetector, 
                self.MeasUnit)[0]

            # deactivate for debugging, activate after debugging
            # get valid status
            meas_status = self.APX.get_valid_status(self.benchDetector)
            assert meas_status == True, "Measurement is not valid" #if measurement is not valid, assertion error TODO: Handle non-valid measurements

            self.APX.OUTPUTGEN_OFF()

            # deactivate for debugging, activate after debugging
        step_result = self.step_passed(
            meas_ch1, 
            self.LimitMin, 
            self.LimitMax)

        print(f"{self.TS_DESC}\tMeas_Val:{meas_ch1:.3f}\tUnit:{self.MeasUnit}\tTol_Min:{self.LimitMin:.3f}\tTol_Max:{self.LimitMax:.3f}\tPassed:{step_result}")
        return meas_ch1, step_result #print or log



    def step_passed(self, meas:float=0, min:float=-1.0, max:float=1.0):
        """Checks test_step status\n
        returns measurement value and status (True:Pass, False:Fail)"""
        assert max >= min, "Limits Error: Min >= Max"
        if (min <= meas <= max): #same as (meas > min) & (meas <= max)
            return True 
        else:
            return False


    def init_PWR_SPLY(self):
        #init values PWR 1-3
        self.IMAX1_A = 0.1 #unit: A
        self.IMAX2_A = 0.5
        self.IMAX3_A = 0.5
        #old assignment
        self.PWR1_VOUT = 56 # variable voltage, unit: V
        self.PWR2_VOUT = 15 # +15V, const
        self.PWR3_VOUT = 15 # -15V, const; only positive values possible

        #Keithley Identifiers: 
        ID_PwrSply1 = 'USB::0x05E6::0x2200::9208388::INSTR' # 46V/56V
        ID_PwrSply2 = 'USB::0x05E6::0x2200::9208769::INSTR' # +15V
        ID_PwrSply3 = 'USB::0x05E6::0x2200::9208692::INSTR' # -15V


        #instances and perform init
        self.pwr1 = Keithley_2200(ID_PwrSply1)
        self.pwr2 = Keithley_2200(ID_PwrSply2)
        self.pwr3 = Keithley_2200(ID_PwrSply3)

        # assignment test, 14.07.2022: OK
        self.pwr1.set_Vout_Imax(self.PWR1_VOUT,self.IMAX1_A)
        self.pwr1.set_output_on()
        self.pwr2.set_Vout_Imax(self.PWR2_VOUT,self.IMAX2_A)
        self.pwr2.set_output_on()
        self.pwr3.set_Vout_Imax(self.PWR3_VOUT,self.IMAX3_A)
        self.pwr3.set_output_on()

        
    def init_DAQ(self):
        #DAQ970A
        self.DAQ = DAQ970A() #Instrument IDs in class DAQ970A
        # Turn Relais Off
        self.DAQ.REL.SPON_All()
        print(self.DAQ.REL.Qery_SysDate()) 


    def initAPx(self):
        # start APx555 Application with project file
        self.load_APxTemplate() #TODO: check default settings of project file
        self.APX.OUTPUTGEN_OFF() # Setting in APx Init
        self.APX.set_XLR_Unbal_Zout(API.AnalogBalancedSourceImpedance.SourceImpedance_40)
        self.APxInput = API.APxInputSelection.Input1 #init: only one input (with two channels)
        self.APxChannel = API.InputChannelIndex.Ch1 #init: only one channel used

        self.APX.set_Zin(
            API.InputConnectorType.AnalogBalanced, 
            API.InputChannelIndex.Ch1, 
            API.AnalogInputTermination.InputTermination_Bal_200k
            )
        SETTLING_TOL_PRC = float(0.1) #in %
        self.APX.set_Settling_Tolerance(SETTLING_TOL_PRC)

        self.LoPassFilterType = API.LowpassFilterModeAnalog.Butterworth
        self.HiPassFilterType = API.LowpassFilterModeAnalog.Butterworth

    def V_to_dBu(self, V_val:float) -> float: 
        '''
        calculate V to dBu, return dBu, all values as float
        '''
        Input_Level = math.sqrt(0.6) #1mW an 600Ohm
        return (20*math.log10(V_val / Input_Level ))

    def dBu_to_V(self, dBu_val:float) -> float: 
        '''
        calculate dBu to V, return V, all values as float
        '''
        Input_Level = math.sqrt(0.6) #1mW an 600Ohm
        
        return (10**(dBu_val / 20) * Input_Level)


    def __del__(self):
        '''
        shutdown teststand
        '''
        self.pwr1.set_output_off()
        self.pwr2.set_output_off()
        self.pwr3.set_output_off()
        self.DAQ.REL.SPON_All()
        self.APX.OUTPUTGEN_OFF()





EOL1493 = SH_1493_Test()
# EOL1493.TS00()
# EOL1493.TS01()
# EOL1493.TS02()
# EOL1493.TS03()
# EOL1493.TS04()
# EOL1493.TS05()
# EOL1493.TS06()
# EOL1493.TS07()
# EOL1493.TS08()
# EOL1493.TS09()
EOL1493.TS10()
EOL1493.TS11()
EOL1493.TS12()
EOL1493.TS13()
EOL1493.TS14()
EOL1493.TS15()
EOL1493.TS16()
EOL1493.TS17()
EOL1493.TS18()
EOL1493.TS19()
EOL1493.TS20()
EOL1493.TS21()
EOL1493.TS22()
EOL1493.TS23()
EOL1493.TS24()
EOL1493.TS25()
EOL1493.TS26()
EOL1493.TS27()




#Some method Tests
#22.07.2022, BL: Added measurement types
# self.MEAS_TYPE = SH_MEAS.get("APX_FUNC") 
# self.exec_test_step() #22.07.2022: APX_FUNC is recognized correctly in exec_test_step
# self.MEAS_TYPE = SH_MEAS.get("AMP_PWR1")
# self.exec_test_step()  #22.07.2022: AMP_PWR1 is recognized correctly in exec_test_step
# self.MEAS_TYPE = SH_MEAS.get("VOLT_DMM_CH1") 
# self.exec_test_step()  #22.07.2022: VOLT_DMM_CH1 is recognized correctly in exec_test_step

# 22.07.2022, BL: Added load_APxTemplate, tests were performed with wrong filename, with wrong timestamp. Exceptions @ fileNotFound, and timestamp-errors are OK. APx loads the correct application @ no exceptions

#21.07.2022: step_passed test, BL, OK
# assert SH_REL.get('REL_Last_1K') == 1, "Error@REL_Last_1K"  #no Error, OK
# assert SH_REL.get('REL_MP106') == 8, "Error@REL_MP106"  #no Error, OK
# assert SH_REL.get('REL_MP102') == 19, "Error@REL_MP102"  #no Error, OK
# assert SH_REL.get('') is None #no Error, OK
# assert SH_REL.get(None) is None #no Error, OK
# try:
#     assert SH_REL.get(None) is int, "None is no int" #crosscheck, exception occurred, OK
# except:
#     print("Assertion error: Generated the Correct Exception")


# translate Relay-Name to its int-Value
# tested, OK, 14.07.2022, BL
# assert SH_REL.get('REL_Last_1K') == 1 #OK
# assert SH_REL.get('REL_MP106') == 8 #OK
# assert SH_REL.get('REL_MP102') == 19 #OK
# assert SH_REL.get('') is None #OK

#turnoff is in init Keithley, OK, BL, 14.07.2022
# self.pwr1.turn_off() 
# self.pwr2.turn_off()
# self.pwr3.turn_off()
