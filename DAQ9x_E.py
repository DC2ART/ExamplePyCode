# Sennheiser Driver Testmodule
# 24.01.2022
# BL
# Addressed Instruments:
# DAQ970A (Subsystem Control, Digital Multimeter aka DMM),
# DAQ903M (20 General Purpose Relays),
# DAQ901A (20 Ch Relais Card Interface to Multimeter)

# Relevant Documentation: 
# Keysight DAQ970A Data Acquisition System: file:///C:/Users/Prozesstechnik/Documents/Sennheiser_Keysight/Keysight_DAQ907A_UserManual.pdf

# Programming Guide Keysight DAQ970A/DAQ973A Data Acquisition System
# file:///C:/Users/Prozesstechnik/Documents/Sennheiser_Keysight/Keysight_DAQ907AInstr_ProgGuide.pdf

# Edit:
# Edit 14.06.2022:  
# Try Keysight Support help(ACQ:VOLT:DC 100,7,0.0001,(@103) --> TRIG:DEL 0.5 --> INIT --> time.sleep(float(self.MEAS_DELAY_SEC)) --> FETC?) : That dident work!!!   
# Timeout creat (self.timeout =20000) : That dident work!!!

# Edit 15.06.2022: 
# clear Timeout, clear query ("TRIG:SOUR BUS") and use Default Trigersource     (IMM) , clear query ("CONF?") : That work and no Error --> OK !!!

# 20.06.2022: TRIG.BUS test: IT DOES NOT WORK!!!! "TRIG:SOUR BUS","TRIG:DEL "+ str(self.MEAS_DELAY_SEC),"INIT" , time.sleep(float(self.MEAS_DELAY_SEC)); *TRG


#from tkinter import TRUE
from tkinter.tix import CheckList
import pyvisa
import time
import logging
import re
logging.basicConfig(filename="Error.log", format="%(asctime)s - %(message)s", level= logging.ERROR)

# DAQM903A: 20 independant SPST Relays in DAQ970A
# Progr Reference: Keysight_DAQ907AInstr_ProgGuide
class DAQ903A:
    def __init__(self, slot_id, base_instr):
        #self.instr = instrument
        self.N_CHANNELS = 20    #Number of channels
        self.SLOT_ID = slot_id
        self.base_instr = base_instr

    def Set_Close(self, channel_list):
        '''
        close non-exclusive: ROUTE:CLOSE close relays in channel list and leave state of all other relays as is
        '''
        
        ch_Lst_corr = self.Check_Ch_Lst(channel_list) #Format Channel List
        resp = self.base_instr.write("ROUTE:CLOSE (@" + ch_Lst_corr  + ")")
        return(resp)
        

    def Set_XClose(self, channel_list):
        '''
        close relays in channel list and open all other relays
        Note 29.07.2022: Command works only for one channel, not a list of channels. With a list of channels only the first channel is closed all other channels are opend. Bug?
        '''
        
        ch_Lst_corr = self.Check_Ch_Lst(channel_list) #Format Channel List
        resp = self.base_instr.write("ROUT:CLOS:EXCL (@" + ch_Lst_corr  + ")")
        return(resp)
        

    # Query channel-list for Close state. If Close == True 
    # -> return-value = 1, else 0
    def Query_Close(self, channel_list:str = "1"):
        ch_Lst_corr = self.Check_Ch_Lst(channel_list) #Format Channel List
        return(self.base_instr.query("ROUTE:CLOSE? (@" + ch_Lst_corr + ")"))
        print("DAQ903A: Channels " + ch_Lst_corr + " Closed (1): " )

    #Open Relays in channel-list
    def Set_Open(self, channel_list:str = "1"):
        ch_Lst_corr = self.Check_Ch_Lst(channel_list) #Format Channel List
        return(self.base_instr.write("ROUTE:OPEN (@" + ch_Lst_corr + ")"))

    # Query channel-list for Open state. If Open == True -> 
    # return-value = 1, else 0
    def Query_Open(self, channel_list):
        #ch_list e.g. "101,107,112"
        #response is e.g. '1,0,1\n'
        ch_Lst_corr = self.Check_Ch_Lst(channel_list) #Format Channel List        
        return(self.base_instr.query("ROUTE:OPEN? (@" + ch_Lst_corr + ")"))
        print("DAQ903A: Channels " + ch_Lst_corr + " Opened (1): " )

    # set relais of a certain slot to Power ON
    def SPON_Ch(self):
        return(self.base_instr.write("SYST:CPON " + str(self.SLOT_ID)))

    # set all channels to Power ON
    def SPON_All(self):
        '''
        Opens all relais in all modules/slots (REL, DMM)
        '''
        return(self.base_instr.write("SYST:CPON ALL"))
    
    # get switch status
    def Sw_Status(self):
        #return value: 1 -> switching finished, 0 unfinished
        return(self.base_instr.query("ROUTE:DONE?"))

    def Qery_SysDate(self):
        # return "+yyyy,+mm,+dd"
        return(self.base_instr.query("SYSTEM:DATE?"))

    # Get cycle count of relays in channel-list as string
    def GetCycleCount(self, channel_list):
        # ch_list e.g. "101,107,112"
        # response is e.g. '+231,+421\n'
        ch_Lst_corr = self.Check_Ch_Lst(channel_list) #Format Chl List        
        return(self.base_instr.query("SYST:REL:CYCL? (@" + ch_Lst_corr + ")"))

    def Check_Ch_Lst(self, channel_Nr:str):
        # All Checks and append actions for all channels in channel_list, e.g. channel_list = "1, 02, 13" -> output channel_list: "201, 202, 213"

        # init return str
        ret_ch_lst = ""
        # create channel list
        channel_lst = channel_Nr.split(",")
        #parse through list
        for i in range(0,len(channel_lst)):
            # eliminate space
            ch = channel_lst[i].replace(" ", "")
            print(ch)
            if len(ch) == 2:
                assert (ch[0]=="0" or ch[0]=="1") & (re.match("\\d",ch[1]) != None), "Check_Ch_Lst: leading zero and integer required in channel list"
                ch = str(self.SLOT_ID) + ch
                ret_ch_lst = ret_ch_lst + "," + ch
            elif len(ch) == 1:
                assert re.match("\\d",ch[0]) != None, "Check_Ch_Lst: integer required in channel list"
                ch = str(self.SLOT_ID) + "0" + ch
                ret_ch_lst = ret_ch_lst + "," + ch
            else:
                print("Check_Ch_Lst: Error, channel_NO must contain only two digit")
        return(ret_ch_lst[1:len(ret_ch_lst)])


#DAQM901A: Relay Interface to Multimeter DAQ970A
#Progr Reference: Keysight_DAQ907AInstr_ProgGuide
class DAQ901A:
    def __init__(self, slot_id, base_instr):
        self.base_instr = base_instr
        self.N_CHANNELS = 20    #Number of channels
        self.SLOT_ID = slot_id
        # Reference to ini-File:
        # MEAS_DURATION_SEC = SAMPLE_TIME_SEC x SAMPLE_COUNT
        # The calling Function sets SAMPLE_COUNT according MEAS_DURATION_SEC initializes SAMPLE_TIME_SEC and leaves SAMPLE_TIME_SEC constant
        self.SAMPLE_TIME_SEC = 0.0001 #Duration for one sample in sec
        self.SAMPLE_COUNT = 7 #Number of Samples taken per measurment, , dummy value

        self.MEAS_DELAY_SEC = 3 # Delay before measurement, dummy value
        # Reference to ini-File: MEAS_DELAY_SEC
        
        self.DMM_VDC_RANGE = "100 V" #Range for DC Measurement in V ----> Mina: Keysight_DAQ907A_UserManual.pdf beginn von seite 109
        self.V_MEAS_RES_V = 0.00001  #--->Mina: ?
        
        self.base_instr.timeout =20000 # 27.06
        

    def err_log(self, qstr, resp):
        err_res = str(self.base_instr.query("SYST:ERR?")).replace('\n','')
        ret_str = ("Query: \"" +  str(qstr) + "\" response: " + str(resp) + " Errors: " + str(err_res))
        return ret_str


    def Meas_Volt_DC(self, channel_Nr:str = "01", MEAS_DUR_SEC:float=0.2, MEAS_DELAY_SEC:float=0.5):
        '''Voltage DC Measurement with DAQ970A.\n
        Allowed measurment channels: 01, 02\n
        Measurement duration in sec,\n
        Measurement delay in sec from testplan from customer
        Input Param. MEAS_DUR_SEC defines SAMPLE_COUNT, SAMPLE_TIME_SEC is fix
        Return parameter is the average measurement value over N samples, N = MEAS_DUR_SEC/SAMPLE_TIME_SEC
        Tested in Testmodule DAQ970.py, result: OK
        '''
        
        self.MEAS_DELAY_SEC = MEAS_DELAY_SEC
        self.SAMPLE_COUNT = int(MEAS_DUR_SEC/self.SAMPLE_TIME_SEC)

        #Info_Prefix = "DAQ901A_Meas_Volt_DC: "
        channel_list = self.Check_Ch_Lst(channel_Nr) 

        
        Info_Prefix = "Meas_Volt_DC: " # for info-messages

        #Configure VDC Measurement
        qstr = "ACQ:VOLT:DC "+str(self.DMM_VDC_RANGE)+","+str(self.SAMPLE_COUNT)+","+str(self.SAMPLE_TIME_SEC)+",(@"+channel_list+")"
        self.base_instr.write(qstr) #No return
        resp = "\"None\"" #No return value
        print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors
        

        #Check ACQ Config: Ch, func, range, sampleCount. sampleTime
        qstr = "ACQ?"
        resp = self.base_instr.query(qstr)
        self.err_log(qstr, resp)# Get and show Errors

        '''
        #Check Trigger Source
        qstr = "TRIG:SOUR?"
        resp = self.base_instr.query(qstr)
        print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors

        #Check Trigger Delay
        qstr = "TRIG:DEL?"
        resp = self.base_instr.query(qstr)
        print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors

        #Set Delay Time
        qstr = "TRIG:DEL " + str(self.MEAS_DELAY_SEC) # Sets the delay between the trigger signal and the first measurement.
        resp = self.base_instr.write(qstr)
        print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors

        #Check Trigger Delay
        qstr = "TRIG:DEL?"
        resp = self.base_instr.query(qstr)
        print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors
        '''

        # Set Trigger from idle into "wait-for-trigger" state
        qstr = "INIT"  #no response
        self.base_instr.write(qstr)
        resp = " \"None\" "
        print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors
        # Trigger Measurement
        
        
        time.sleep(float(self.MEAS_DELAY_SEC))

        qstr = "FETC?" #READ?
        resp = self.base_instr.query(qstr)
        # print(Info_Prefix + self.err_log(qstr, resp) ) # Get and show Errors

        
        Meas_List = resp.split(",")
        for i in range(0, len(Meas_List)):
            Meas_List[i] = float(Meas_List[i])
        Meas_AVG = sum(Meas_List)/float(len(Meas_List))

        return Meas_AVG

        # check Temperature

    

    def meas_temp(self):
        qstr = "TEMPerature:RJUNction?"
        resp = self.base_instr.query(qstr)
        

        return resp

        #Optional: Save Samples in File in defined cases e.g. Pass, Fail, Golden Sample to analyse Voltage Reading
    

    def Check_Ch_Lst(self, channel_Nr:str):
        # All Checks and append actions for all channels in channel_list, e.g. channel_list = "1, 02, 13" -> output channel_list: "201, 202, 213"

        # init return str
        ret_ch_lst = ""
        # create channel list
        channel_lst = channel_Nr.split(",")
        #parse through list
        for i in range(0,len(channel_lst)):
            # eliminate space
            ch = channel_lst[i].replace(" ", "")
            if len(ch) == 2:
                assert (ch[0]=="0" or ch[0]=="1") & (re.match("\\d",ch[1]) != None), "Check_Ch_Lst: leading zero and integer required in channel list"
                ch = str(self.SLOT_ID) + ch
                ret_ch_lst = ret_ch_lst + "," + ch
            elif len(ch) == 1:
                assert re.match("\\d",ch[0]) != None, "Check_Ch_Lst: integer required in channel list"
                ch = str(self.SLOT_ID) + "0" + ch
                ret_ch_lst = ret_ch_lst + "," + ch
            else:
                print("Check_Ch_Lst: Error, channel_NO must contain only two digit")
        return(ret_ch_lst[1:len(ret_ch_lst)])


# DAQ970A: Mulitmeter Instrument
class DAQ970A:
    def __init__(self):
        
        # Instrument ID
        self.IDStr_970A = "USB::0x2A8D::0x5101::MY58010326::INSTR"
        # Slot IDs for query and assignment
        self.IDStr_903A = "Keysight Technologies,DAQM903A,MY58002667,01.02"
        self.IDStr_901A = "Keysight Technologies,DAQM901A,MY58022244,01.02"
        # query response for empty slots
        self.IDStr_Empty = "Keysight Technologies,0,0,0" 
        self.N_SLOTS = 3 #Maximum number of slots
        Info_Prefix = "DAQ970A INIT: "
        #self.instr = pyvisa.ResourceManager().open_resource(self.IDStr_970A)
        #TODO: Error Handling for opening Source e.g. Instrument not present, wrong ID ... ---------------->Done
        try:
           self.instr = pyvisa.ResourceManager().open_resource(self.IDStr_970A)
        except pyvisa.errors.VisaIOError as err:
            print(Info_Prefix + str(err))
            logging.error(err)
            rm = pyvisa.ResourceManager()
            lst = rm.list_resources('?*')
            print(lst)
        except Exception as err:
            print(err)
        else:
            print(Info_Prefix +"Instrument is present")
            


        
        Info_Prefix = "DAQ970A INIT: " # for info-messages
        
        #self.instr.timeout =10000

        # Get Errors
        qstr = str(self.GetErrors()).replace('\n','')
        print(Info_Prefix + "System Errors: " +  qstr )
        # TODO: Check if Errors need notification etc.->Error Handling

        # Clear Errors
        qstr = str(self.ClearStatus())
        print(Info_Prefix + "Clearing Errors :" + qstr)

        # Get ID
        qstr = '*IDN?'
        resp = self.instr.query(qstr)
        print(Info_Prefix + "IDN: " + resp)
        # TODO: Check expected ID self.IDStr_970A vs. read qstr->Error Handling ---------------->DONE
        
        IDN_SR = "MY58010326"
        if IDN_SR in resp and self.IDStr_970A:
            print(Info_Prefix +"instrument IDs is Correct")
        else:
            print(Info_Prefix +"Warning! wrong instrument is connected ")

        #Reset Unit
        print(Info_Prefix + "Resetting Instrument to Factory Settings ") 
        self.ResetInstr() #no return value

        # def Set_SysDate, for timestamp of measurement data
        # TODO: implement this -->Datum von rechner
        
 
        # Get System-Time
        qstr = "SYSTem:TIME?"
        resp = self.instr.query(qstr)
        print(Info_Prefix + "SYSTem:TIME: " + resp)

       
        # Perform Self Test
        # self.SelfTest()
        # 09.06.2022: Exception has occurred: VisaIOError
        # TODO: Correct this Error------------------------>Done
        # TODO: Selftest: Errorhandling ------------------>Done
        ''''''
        try:
            qstr = str(self.SelfTest())
        except  pyvisa.errors.VisaIOError as err:
            print(Info_Prefix + str(err))
        else:
            if int(qstr) == 0:
                print(Info_Prefix + "SelfTest: Pass")
            elif int(qstr) > 0:
                print(Info_Prefix + qstr)
        

        #Get DMM-Slot
        self.Slot_Assign()
        print(Info_Prefix + f"Slot-# DAQ901A is: {self.Slot_No_DMM}")
        print(Info_Prefix + f"Slot-# DAQ903A is: {self.Slot_No_REL}")

        # TODO: Temperature measurement and monitoring for overheating, permissible range: initially 15°C ... 28°C --------------------> Done   MEASure:TEMPerature:<probe_type>? Seite 174 , [SENSe:]TEMPerature Subsystem seite 256

        # Check if Errors occurred during init
        qstr = str(self.GetErrors()).replace('\n','')
        print(Info_Prefix + "System Errors: " +  qstr )
        # TODO: Check if Errors need notification etc.->Error Handling

        # Clear Errors
        qstr = str(self.ClearStatus())
        print(Info_Prefix + "Clearing Errors :" + qstr)

    def GetID(self):
        return(self.instr.query("*IDN?"))
    def GetErrors(self):
        return(self.instr.query("SYST:ERR?"))
    def ClearStatus(self):
        return(self.instr.write("*CLS"))
    def SelfTest(self):
        return 0 #to be implemented
        # return(self.instr.query("TEST:ALL?"))
        
    def ResetInstr(self):
        return(self.instr.write("*RST"))

    #Identify instruments in slots, instruments occur only onces
    def Slot_Assign(self):
        for slot in range(1, self.N_SLOTS + 1 ):
            resp = self.instr.query("SYSTem:CTYPe? " + str(slot))
            if(resp.find(self.IDStr_903A)==0):
                self.Slot_No_REL = slot
                self.REL = DAQ903A(slot, self.instr)
            elif(resp.find(self.IDStr_901A)==0):
                self.Slot_No_DMM = slot
                self.DMM = DAQ901A(slot, self.instr)
            elif(resp.find(self.IDStr_Empty)==0):
                print(f"Info DAQ970A: Slot {slot} ist empty")
            if (slot >= self.N_SLOTS and self.DMM is None):
                print("DAQ970A Error: Slot_Assign: DAQ901A not fount")
                #TODO: ErrorHandling e.g. Operator Notification, Beep ... 
            if (slot >= self.N_SLOTS and self.REL is None):
                print("DAQ970A Error: Slot_Assign: DAQ903A not fount")
                #TODO: ErrorHandling e.g. Operator Notification, Beep ... 
