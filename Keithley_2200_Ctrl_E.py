# Keithley 2200-60-2.5 Powersupply Controll
# Vorlage von BL: Keithley_2200_Ctrl.py 
# Class: Keithley_2200, methods implemented and tested: __init__, query_str, write, identify, reset_instrument, get_errors, clear_errors, set_output_on, set_output_off, set_Vout_Imax, turn_off, meas_volt, meas_curr
# Reference: Keithley Series 2200 Multichannel Programmable DC Power Supplies, # Programming Reference
#"C:\Users\Prozesstechnik\Documents\Sennheiser_Keithley\PDF\2200S-901-01_D_May_2017_Ref.pdf"

# BL
# 21.06.2022
# Update:
# Init Strings of Netzteil 1 Keithley 2200-60-2 was not correct and belongs to Netzteil 3 -----> EOL_SH_Init.py must to be corrected
# Init Strings of Netzteil 4 Keithley 2200-60-2 belong to Netzteil 1

from os import stat
import time
import pyvisa

class Keithley_2200():
    def __init__(self, in_resource):
    
        self.resource = in_resource
        self.instr = pyvisa.ResourceManager().open_resource(self.resource)
        self.write("*OPC") #set OPC flag to enable "done" status: *OPC?
        response = self.identify()
        #print(response)
        self.PWR_SETTLING_TIM_SEC = 0.5 #Time Keithly 2200 need to settle the output parameters
        if 'Keithley' not in response:
            raise Exception ('wrong instrument! expected Keithley but got ' + response)
       
        #print("Note: Voltage Limit: " + self.instr.query("VOLTAGE:LIMIT?") +"V")


        self.instr.write("*RST") #Reset
        self.instr.write("SYST:REMOTE") #Set to Remote-Mode
        # set locked out front panel, This command has no effect if the
        # instrument is in local mode.
        self.instr.write("SYSTem:RWLock")
        

        # set Trigger source
        self.instr.write("TRIGER:SOURCE: BUS")
        resp = self.instr.query("TRIGger:SOURce?")
        #print (resp)

        
        # set OVP
        self.instr.write("VOLT:PROT 60V")
        self.instr.write("VOLT:PROT:STAT ON")
        Stat = self.instr.query("VOLT:PROT:STAT?")
        #print ("stat:" + str(Stat))
        # might return 1, which would indicate that 
        #overvoltage protection is active.
        self.turn_off()

        # query System Errors
        def Error(self):
            Error = self.instr.query("SYSTem:ERRor?")
            res_of_err = ("Error: " + str(Error))
            return res_of_err
        
    def __del__(self):
        self.instr.close()

    def query_str(self, cmd: str):
        """wrapper for instrument query"""
        return(self.instr.query(cmd)).replace("\n", "")

    def write(self, cmd: str):
        """wrapper for instrument write"""
        self.instr.write(cmd)

    def identify(self):
        return(self.query_str('*IDN?'))

    def reset_instrument(self):
        self.write("*RST")

    def get_errors(self):
        return(self.query_str("SYST:ERR?"))

    def clear_errors(self):
        self.write("*CLS")

    def set_Vout_Imax(self, voltage, current):
        voltage = float(round(voltage,5))
        current = float(round(current,5))
        self.write("VOLTAGE " + str(voltage))
        self.write("CURRENT " + str(current))
        time.sleep(self.PWR_SETTLING_TIM_SEC)

    def set_output_on(self):
        self.write("SOURCE:OUTPUT:STATE ON")
        #self.write("OUTP 1")

    #set only output off, V, I values are not affected
    #14.07.2022: Tested, OK, BL
    def set_output_off(self):
        self.write("SOURCE:OUTPUT:STATE OFF")
        #self.write("OUTP 0")

    def turn_off(self):
        self.set_Vout_Imax(0, 0)
        self.write("OUTP 0")

    def meas_volt(self, delay_sec:float=0.0, duration_sec:float=0.0):
        '''Voltage (DC) measurement ( in V)\n
        return value is the average current in A over N measurements\n
        Number of samples is duration_sec / approx. 20ms (sample time) 
        Wait for delay in sec (python implementation, no real time clock)
        '''
        q_str = "MEASURE:VOLTAGE:DC?"
        return self.meas_frame(delay_sec, duration_sec, q_str)

    def meas_curr(self, delay_sec:float=0.0, duration_sec:float=0.0):
        '''Current (DC) measurement ( in A)\n
        return value is the average current in A over N measurements\n
        Number of samples is duration_sec / approx. 20ms (sample time) 
        Wait for delay in sec (python implementation, no real time clock)
        '''
        q_str = "MEASURE:CURRENT:DC?"
        return self.meas_frame(delay_sec, duration_sec, q_str)

    def meas_frame(self, delay_sec, duration_sec, q_str):
        '''handles delay, duration, average-calculation, status-check\n
        Functionalities common to voltage and current measurement
        '''
        # initialize array with measurement values
        meas_array = []
        
        # wait delay seconds
        time.sleep(delay_sec)
        
        # calculate end time
        t_end = time.time() + duration_sec
        
        # measure (recursively) for duration seconds
        while time.time() < t_end:
            OPC = self.query_str("*OPC?") #1: done, 0:busy
            # print(f"TimeNow:{time.time()}\\OPCStatus:{OPC}\\Number of Samples: {len(meas_array)}") #for debug optionally for logging
            if OPC == "1":
                meas_array.append( float(self.query_str(q_str).replace("\n", "")) ) # append result to meas_array
            # self.check_status() #integrate status-check into measurement flow: TODO
        
        print(f"measurement values: {meas_array}") #for logging / debug
        return sum(meas_array)/len(meas_array) # return average measurement 

    def check_status(self):
        # Check Status, e.g. SBR, SESR, OCR, QEVR: TODO
        # status = ... get register contents, evaluate bits
        # return status # status = 0 if OK, status = nonZero if not OK, with Error codes from Keithly e.g. -220: Execution error, 40: Flash Write failed etc.

        pass 
