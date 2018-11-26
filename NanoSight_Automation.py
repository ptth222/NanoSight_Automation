import wx
import threading
from pywinauto.application import Application
import pywinauto
import time
import re
import serial
import serial.tools.list_ports
import pandas
import os
import uiautomation as automation
import errno
import wx.lib.agw.genericmessagedialog as GMD

__version__ = "1.0"

## TODO In case user forgets to uncheck the date and timestamp look to see if the nan file 
## exists eith the expected name and if not then look for one with a datestamp and load that 
## one to process instead.

EVT_PULLED_PLUG = wx.NewId()
EVT_THREAD_ABORTED = wx.NewId()

## Make new events so the thread can signal the main GUI.
class PulledPlugEvent(wx.PyEvent):
    def __init__(self):
        wx.PyEvent.__init__(self)
        self.SetEventType(EVT_PULLED_PLUG)
        



class ThreadAbortedEvent(wx.PyEvent):
    def __init__(self):
        wx.PyEvent.__init__(self)
        self.SetEventType(EVT_THREAD_ABORTED)
        

        
        

## Class taken from https://wiki.wxpython.org/LongRunningTasks and modified.
# Thread class that executes processing
class BatchThread(threading.Thread):
    """Batch Thread Class."""
    def __init__(self, batch_data):
        """Init Serial Thread Class."""
        threading.Thread.__init__(self)
        self.batch_data = batch_data
        self.sample_df = batch_data.sample_df
        self.want_abort = False
        self.daemon = True
        # This starts the thread running on creation, but you could
        # also make the GUI thread responsible for calling this
        self.start() 
    
    
    def run(self):
        """Run Batch Thread."""
        
        ## Create save directories.
        if not self.Create_Save_Directories():
            wx.PostEvent(self.batch_data, ThreadAbortedEvent())
            return
        
        
        ## Connect to Arduino, NTA, and CETAC.
        Arduino_connect_result = self.Connect_To_Arduino()
        if not Arduino_connect_result or not self.Connect_To_NTA() or not self.Connect_To_CETAC():
            if Arduino_connect_result:
                self.ComPort.close()
            wx.PostEvent(self.batch_data, ThreadAbortedEvent())
            return
        
        
        ## Reset output latch in Arduino.
        ## This reset of the latch is to make sure the latch starts from a known state before the batch begins.
        while True:
            send_response = self.Reset_AS_Output_Latch()
            
            if send_response == "Latch Cleared":
                break
            
            elif send_response == "Abort":
                self.Abort_Clean_Up()
                return
            
            elif send_response == "Pulled Plug":
                self.PP_Clean_Up()
                return
            
            elif send_response == "Time Out":
                message = "Time out reached while trying to communicate with the Arduino. \nCheck that the Arduino is functioning properly. \nClick Yes to try sending the signal again."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.YES_NO | wx.ICON_QUESTION)
                answer = msg_dlg.ShowModal()
                msg_dlg.Destroy()
                if not answer == wx.ID_YES:
                    self.Abort_Clean_Up()
                    return

        
        ## Make sure CETAC is connected to the autosampler.
        if not self.CETAC_Check_For_COM():
            self.ComPort.close()
            wx.PostEvent(self.batch_data, ThreadAbortedEvent())
            return
        
        ## Make sure CETAC has a script loaded.
        if not self.CETAC_Check_For_Active_Script():
            self.ComPort.close()
            wx.PostEvent(self.batch_data, ThreadAbortedEvent())
            return
        
        ######################
        ## Acquire the samples.
        ######################
        for i in range(len(self.sample_df)):
            
            ## Before starting make sure the user hasn't already wanted to abort.
            if self.want_abort:
                self.Abort_Clean_Up()
                return
            
            ## Set the label in acquistion progress to in progress and change color to blue.
            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "In Progress")
            self.batch_data.sample_list_ctrl.SetItemBackgroundColour(i, "light blue")
            
            ## Set base filename for sample.
            if not self.NTA_Set_Filename(self.sample_df.loc[:, "Save Directory"][i], self.sample_df.loc[:, "Sample Name"][i]):
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            ## Set up correct script in autosampler and NanoSight.
            if not self.NTA_Load_Script(self.sample_df.loc[:, "Acquire Script"][i]):
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            
            ## Check for abort again.
            if self.want_abort:
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            

            
            ## Flush the input buffer before starting the autosampler script to make sure
            ## signals from a previous run or false signals aren't misinterpreted.
            self.ComPort.reset_input_buffer()
            self.ComPort.flushInput()

            
            ## If this is the first sample start the script, otherwise send trigger to MVX if it is ready.
            if i == 0:
                
                ## Close any windows that may be left open by the user.
                self.CETAC_Close_All_Windows()
            
                ## Start autosampler script.
                run_response = self.CETAC_Run_Script()
                if run_response == "Run Script Button Disabled" or run_response == "Abort Script Button Disabled" or run_response == "No Samples Selected":
                    self.ComPort.close()
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    wx.PostEvent(self.batch_data, ThreadAbortedEvent())
                    return
                
                ## Check for errors after starting the script.
                if self.CETAC_Check_For_Errors():
                    self.ComPort.close()
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    wx.PostEvent(self.batch_data, ThreadAbortedEvent())
                    return
                

            else:    
            
                ## Check for errors.
                if self.CETAC_Check_For_Errors():
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    self.Abort_Clean_Up()
                    return
                
                ## Wait for signal from autosampler.
                timeout = 15
                listen_response = self.Check_AS_Output_Latch(timeout)
                ## Check to see if the batch was aborted while listening or the Arduino was disconnected.
                if listen_response == "Abort" or listen_response == "Time Out":
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    self.Abort_Clean_Up()
                    return
                
                elif listen_response == "Pulled Plug":
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    self.PP_Clean_Up()
                    return
                
                
                
                ## Wait some time before sending signal to autosampler to make sure 
                ## it has had time to move on to the trigger command, and so we do 
                ## not relatch the same output signal twice.
                time.sleep(6)
                
                ## Reset output latch in Arduino.
                ## If something goes wrong with the Arduino at this point then it can
                ## be fixed or replaced and the batch can continue so allow the user to retry sending the signal to the Arduino.
                while True:
                    send_response = self.Reset_AS_Output_Latch()
                    
                    if send_response == "Latch Cleared":
                        break
                    
                    elif send_response == "Abort":
                        self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                        self.Abort_Clean_Up()
                        return
                    
                    elif send_response == "Pulled Plug":
                        self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                        self.PP_Clean_Up()
                        return
                    
                    elif send_response == "Time Out":
                        message = "Time out reached while trying to communicate with the Arduino. \nCheck that the Arduino is functioning properly. \nClick Yes to try sending the signal again."
                        msg_dlg = wx.MessageDialog(None, message, "Warning", wx.YES_NO | wx.ICON_QUESTION)
                        answer = msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        if not answer == wx.ID_YES:
                            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                            self.Abort_Clean_Up()
                            return
                        
           
            
            
                ## Send signal to autosampler.
                ## If something goes wrong with the Arduino at this point then it can
                ## be fixed or replaced and the batch can continue so allow the user to retry sending the signal to the AS.
                while True:
                    send_response = self.Send_Signal_To_AS()
                    
                    if send_response == "Signal Sent":
                        break
                    
                    elif send_response == "Abort":
                        self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                        self.Abort_Clean_Up()
                        return
                    
                    elif send_response == "Pulled Plug":
                        self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                        self.PP_Clean_Up()
                        return
                    
                    elif send_response == "Time Out":
                        message = "Time out reached while trying to signal the autosampler. \nCheck that the Arduino is functioning properly. \nClick Yes to try sending the signal again."
                        msg_dlg = wx.MessageDialog(None, message, "Warning", wx.YES_NO | wx.ICON_QUESTION)
                        answer = msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        if not answer == wx.ID_YES:
                            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                            self.Abort_Clean_Up()
                            return
            
            
        
            ## Wait for signal from autosampler.
            timeout = 15
            listen_response = self.Check_AS_Output_Latch(timeout)
            ## Check to see if the batch was aborted while listening or the Arduino was disconnected.
            if listen_response == "Abort" or listen_response == "Time Out":
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            elif listen_response == "Pulled Plug":
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.PP_Clean_Up()
                return
            
            
        
        
            ## Start NanoSight script.
            if not self.NTA_Run_Script():
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.PP_Clean_Up()
                return
            
            
            ## Reset output latch in Arduino.
            ## If something goes wrong with the Arduino at this point then it can
            ## be fixed or replaced and the batch can continue so allow the user to retry sending the signal to the Arduino.
            ## Wait for 5 seconds before resetting latch to make sure the autosampler has turned off its output.
            time.sleep(5)
            while True:
                send_response = self.Reset_AS_Output_Latch()
                
                if send_response == "Latch Cleared":
                    break
                
                elif send_response == "Abort":
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    self.Abort_Clean_Up()
                    return
                
                elif send_response == "Pulled Plug":
                    self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                    self.PP_Clean_Up()
                    return
                
                elif send_response == "Time Out":
                    message = "Time out reached while trying to communicate with the Arduino. \nCheck that the Arduino is functioning properly. \nClick Yes to try sending the signal again."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.YES_NO | wx.ICON_QUESTION)
                    answer = msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    if not answer == wx.ID_YES:
                        self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                        self.Abort_Clean_Up()
                        return
            
            

            ## Wait for end of script signal from NanoSight.
            listen_response = self.Listen_For_End_Of_Script_NTA_Signal()
            ## Check to see if the batch was aborted while listening.
            if listen_response == "Aborted":
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            
            
            ## Set the label in acquistion progress to Complete and change color back to normal.
            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Complete")
            self.batch_data.sample_list_ctrl.SetItemBackgroundColour(i, "white")

        
        ## Give time for NTA to finish completing the script.
        time.sleep(5)
            
            
            
            
            
        
        
        
        ######################
        ## Process the samples.
        ######################
        for i in range(len(self.sample_df)):
            
            if self.want_abort:
                wx.PostEvent(self.batch_data, ThreadAbortedEvent())
                return
            
            ## Set the label in processing progress to in progress and change color to blue.
            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "In Progress")
            self.batch_data.sample_list_ctrl.SetItemBackgroundColour(i, "light blue")
            
            ## Once when running 10 standards in a row after the 4th standard was processed the NTA_Check_Existence
            ## returned False. I am pretty sure it was the check existence call in NTA_Load_Script because the 
            ## program was still in the analysis tab. I think that maybe after exporting there needs to be some 
            ## time before checking existence, so adding the below timer. It should be noted that after the 
            ## existence check failed the NTA program was pretty much locked up, so it might have just been a fluke.
            time.sleep(10)
            ## Turns out it might be that the NTA program can't process more than 4 samples in a row.
            ## When I ran 4 processings in a row manually the same bug happened.
            
            ## Load Process Script
            if not self.NTA_Load_Script(self.sample_df.loc[:, "Process Script"][i]):
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            ## Open experiment to process.
            if not self.NTA_Open_Experiment(self.sample_df.loc[:, "Save Directory"][i], self.sample_df.loc[:, "Sample Name"][i]):
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
                        
            
            if self.want_abort:
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            
            
            ## Start NanoSight script.
            if not self.NTA_Run_Script():
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            ## The assumption is that the process script will use the PROCESSBASIC command

            ## Wait for end of script signal from NanoSight.
            listen_response = self.Listen_For_End_Of_Script_NTA_Signal()
            ## Check to see if the batch was aborted while listening.
            if listen_response == "Aborted":
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            
            ## Export results.
            if not self.NTA_Export_Results():
                self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Cancelled")
                self.Abort_Clean_Up()
                return
            
            
            ## Set the label in acquistion progress to Complete and change color back to normal.
            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Processing Progress"), "Complete")
            self.batch_data.sample_list_ctrl.SetItemBackgroundColour(i, "white")
            
            
            
            
            
            
            
            
            
        ## Give the autosampler the last trigger command so it can end its program.    
        ## Wait for signal from autosampler.
        timeout = 15
        listen_response = self.Check_AS_Output_Latch(timeout)
        ## Check to see if the batch was aborted while listening or the Arduino was disconnected.
        if listen_response == "Abort" or listen_response == "Time Out":
            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
            self.Abort_Clean_Up()
            return
        
        elif listen_response == "Pulled Plug":
            self.batch_data.sample_list_ctrl.SetItem(i, self.batch_data.list_ctrl_col_names.index("Acquisition Progress"), "Cancelled")
            self.PP_Clean_Up()
            return
        
        
        
        ## Wait some time before sending signal to autosampler to make sure 
        ## it has had time to move on to the trigger command.
        time.sleep(6)
        
    
        ## Send signal to autosampler.
        ## If something goes wrong with the Arduino at this point then it can
        ## be fixed or replaced and the batch can continue so allow the user to retry sending the signal to the AS.
        while True:
            send_response = self.Send_Signal_To_AS()
            
            if send_response == "Signal Sent":
                break
            
            elif send_response == "Abort":
                self.Abort_Clean_Up()
                return
            
            elif send_response == "Pulled Plug":
                self.PP_Clean_Up()
                return
            
            elif send_response == "Time Out":
                message = "Time out reached while trying to signal the autosampler. \nCheck that the Arduino is functioning properly. \nClick Yes to try sending the signal again."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.YES_NO | wx.ICON_QUESTION)
                answer = msg_dlg.ShowModal()
                msg_dlg.Destroy()
                if not answer == wx.ID_YES:
                    self.Abort_Clean_Up()
                    return
                
                
        self.ComPort.close()
        
        
        message = "Batch Complete!"
        msg_dlg = wx.MessageDialog(None, message, "Batch Complete", wx.OK)
        answer = msg_dlg.ShowModal()
        msg_dlg.Destroy()
        
        wx.PostEvent(self.batch_data, ThreadAbortedEvent())
        

    
    
    def abort(self):
        """Abort Batch Thread."""
        # Method for use by main thread to signal an abort.
        self.want_abort = True 




    def Abort_Clean_Up(self):
        """Calls the functions to abort the NTA and CETAC scripts, closes the 
        serial connection to the Arduino, and posts the abort event to the main thread."""
        
        self.NTA_Abort_Script()
        self.CETAC_Abort_Script()
        self.ComPort.close()
        wx.PostEvent(self.batch_data, ThreadAbortedEvent())
    
    
    
    
    def PP_Clean_Up(self):
        """Calls the functions to abort the NTA and CETAC scripts, closes the 
        serial connection to the Arduino, and posts the pulled plug event to the main thread."""
        
        self.NTA_Abort_Script()
        self.CETAC_Abort_Script()
        self.ComPort.close()
        wx.PostEvent(self.batch_data, PulledPlugEvent())
    
    
    
    def Create_Save_Directories(self):
        """Create the directories to save the acquisitons and analyses in."""
        
        sample_df = self.sample_df
        
        problem_samples = []
        
        for i in range(len(sample_df)):
            directory = sample_df.loc[i, "Save Directory"]
            sample_name = sample_df.loc[i, "Sample Name"]
            
            if self.batch_data.samples_have_individual_directories:
                directory = os.path.join(directory, sample_name)
                
            try:
                os.makedirs(directory)
            except OSError as e:
                if e.errno != errno.EEXIST:
                    problem_samples.append(sample_name)
            
            
        if len(problem_samples) > 0:
            message = "An error occurred while creating the directories for the following sample(s): \n\n" + "\n".join(problem_samples) 
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        else:
            return True
                    
    
    
    
    
    def Connect_To_Arduino(self):
        """Finds an Arduino in the list of serial ports and connects to it. If there are 
        multiple Arduino's a message is created and no connection is made. If there are no 
        Arduino's a message is created and no connection is made. The serial settings are:
            
            baudrate = 115200
            bytesize = 8
            parity = serial.PARITY_NONE
            stopbits = 1
            timeout = 1"""
        
        ports = [tuple(p) for p in list(serial.tools.list_ports_windows.comports())]
        ports = {name:description for name, description, value2 in ports}
        arduino_port = [name for name, description in ports.items() if re.search(r"Arduino", description)]
        
        if len(arduino_port) == 0:
            message = "Could not find Arduino. Please connect the Arduino to run the batch."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        elif len(arduino_port) > 1:
            message = "There are too many Arduino's connected. The correct one can not be determined."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
            
        else:
            try:
                ComPort = serial.Serial(arduino_port[0])
                ComPort.baudrate = 115200
                ComPort.bytesize = 8
                ComPort.parity = serial.PARITY_NONE
                ComPort.stopbits = 1
                ComPort.timeout = 1
                
                self.ComPort = ComPort
                return True

            except serial.serialutil.SerialException:
                message = "Could not establish connection to Arduino. Check it and try again."
                msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return False
    
    
    
    
    
    def Connect_To_NTA(self):
        """Looks for a program with \"NTA 3.3\" in the title and connects to it using 
        the pywinauto library. If no program is found a message is created."""
    
        windows = pywinauto.findwindows.find_elements()
        if not any([re.search(r"NTA 3.3", str(window)) for window in windows]):
            message = "The NTA 3.3 program is not started. Please start the program and try again."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
            
        else:
            self.NTA_app = Application().connect(title_re = u".*NTA 3.3.*")
            return True
    
    
    
    
    
    def Connect_To_CETAC(self):
        """Looks for a program with \"CETAC Workstation\" in the title and connects to it using 
        the automation library. If no program is found a message is created."""
        
        windows = pywinauto.findwindows.find_elements()
        if not any([re.search(r"CETAC Workstation", str(window)) for window in windows]):
            message = "The CETAC Workstation program is not started. Please start the program and try again."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
            
        else:    
            self.CETAC_app = automation.WindowControl(Name = "CETAC Workstation")
            return True
    
    
    
    
    
    def NTA_Check_Existence(self):
        """Checks to make sure the NTA 3.3 program is still open. Returns True if it is, False if not."""
        
        windows = [str(window) for window in pywinauto.findwindows.find_elements()]
        if not any([re.search(r"NTA 3.3", str(window)) for window in windows]):
            message = "The NTA 3.3 program is not open."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        else:
            return True
        
        
        
    
    
    
    
    def NTA_Set_Filename(self, save_directory, sample_name):
        """Sets the base filename in the NTA 3.3 program according to save_directory and sample_name.
        If the samples have individual directories the path is save_directory\\sample_name\\sample_name.
        Otherwise it is save_directory\\sample_name."""
        
        if not self.NTA_Check_Existence():
            return False
        
        if self.batch_data.samples_have_individual_directories:
            file_path = os.path.join(save_directory, sample_name, sample_name)
        else:
            file_path = os.path.join(save_directory, sample_name)
        
        nta = self.NTA_app["NTA 3.3"]
        ## Wait for the window to be ready.
        time.sleep(2)
        ## Select the tab with the load script button.
        nta.TabControl2.Select(u'SOP')
        time.sleep(2)
        
        ## Check that tab was selected.
        if nta.TabContol2.get_selected_tab() != 0:
            message = "Could not select the SOP tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        ## Select "Recent Measurements" from the combo box.
        nta.TabControl3.Select(u"Recent Measurements")
        time.sleep(2)
        
        ## Check that the combo box option was selected.
        if nta.TableControl3.get_selected_tab() != 0:
            message = "Could not select the Recent Measurements tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        ## Click the load script button.
        nta["..."].click()
        time.sleep(2)
        
        ## Make sure the click was successful and a save dialog was created.
        window_check_result = self.NTA_Window_Check("\'Save As\'", True, 0.2)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            message = "Could not click the \"...\" button to select a base filename in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
            
        
        ## File dialog interaction.
        window = self.NTA_app["Save As"]
        time.sleep(2)
        window["Edit"].set_text(file_path)
        time.sleep(2)
        
        ## Check that the text was edited.
        if window["Edit"].texts()[0] != file_path:
            message = "Could not enter the file path into the Save As dialog in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        window = automation.WindowControl(Name = "Save As")
        window.SetActive(waitTime=1)
        window.ButtonControl(Name = "Save").Click()
        time.sleep(3)
        
        ## Make sure the window closed after clicking open.
        ## If it did not then something went wrong.
        window_check_result = self.NTA_Window_Check("\'Save As\'", False, 0.2)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            message = "Could not click the Save button in the Save As dialog in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        return True
    
    
    
    
    def NTA_Load_Script(self, script_filepath):
        """Loads the script indicated by script_filepath in the NTA 3.3 program."""
        
        if not self.NTA_Check_Existence():
            return False
        
        nta = self.NTA_app["NTA 3.3"]
        ## Wait for the window to be ready.
        time.sleep(2)
        ## Select the tab with the load script button.
        nta.TabControl2.Select(u'SOP')
        time.sleep(2)
        
        ## Check that tab was selected.
        if nta.TabContol2.get_selected_tab() != 0:
            message = "Could not select the SOP tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Load Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        ## Select "Recent Measurements" from the combo box.
        nta.TabControl3.Select(u"Recent Measurements")
        time.sleep(2)
        
        ## Check that the combo box option was selected.
        if nta.TableControl3.get_selected_tab() != 0:
            message = "Could not select the Recent Measurements tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        ## Click the load script button.
        nta["Load Script"].click()
        time.sleep(2)
        
        ## Make sure the click was successful and an Open dialog was created.
        window_check_result = self.NTA_Window_Check("\'Open\'", True, 0.2)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            message = "Could not load a script in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Load Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        
        
        ## File dialog interaction.
        window = self.NTA_app["Open"]
        time.sleep(2)
        window["Edit"].set_text(script_filepath)
        time.sleep(2)
        
        ## Check that the text was edited.
        if window["Edit"].texts()[0] != script_filepath:
            message = "Could not enter the file path into the Open dialog in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Load Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        window = automation.WindowControl(Name = "Open")
        window.SetActive(waitTime=1)
        window.SplitButtonControl(Name = "Open").Click()
        time.sleep(3)
        
        ## Make sure the window closed after clicking open.
        ## If it did not then something went wrong.
        window_check_result = self.NTA_Window_Check("\'Open\'", False, 0.2)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            message = "Could not load a script in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Load Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        return True
    
    


    def NTA_Run_Script(self):
        """Clicks the Run button to run a script in the NTA 3.3 program. This is 
        the button located in the SOP tab under Recent Measurements, not to be 
        confused with the button on the script panel. If a Warning dialog appears 
        after clicking run it is assumed that it is the superfluous one asking about 
        overwriting files and clicks Yes."""
        
        if not self.NTA_Check_Existence():
            return False
        
        nta = self.NTA_app["NTA 3.3"]
        ## Wait for the window to be ready.
        time.sleep(2)
        ## Select the tab with the load script button.
        nta.TabControl2.Select(u'SOP')
        time.sleep(2)
        
        ## Check that tab was selected.
        if nta.TabContol2.get_selected_tab() != 0:
            message = "Could not select the SOP tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Run Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        ## Select "Recent Measurements" from the combo box.
        nta.TabControl3.Select(u"Recent Measurements")
        time.sleep(2)
        
        ## Check that the combo box option was selected.
        if nta.TableControl3.get_selected_tab() != 0:
            message = "Could not select the Recent Measurements tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Set File Name Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        ## Click the run button.
        nta["Run"].click()
        time.sleep(2)
        ## At time of writing there is no way to confirm that the script is running.
        
        ## Close any warning message that may appear such as from the camerasettingmessage command.
        window_check_result = self.NTA_Window_Check("\'Warning\'", False, 0.2)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            window = automation.WindowControl(Name = "Warning")
            window.SetActive(waitTime=1)
            window.ButtonControl(Name = "Yes").Click()
            time.sleep(2)
            
            window_check_result = self.NTA_Window_Check("\'Warning\'", False, 0.2)
        
            if window_check_result == "Abort":
                return False
            elif window_check_result == "Timed Out":
                message = "Could not close the warning dialog in NTA. Batch aborted."
                msg_dlg = wx.MessageDialog(None, message, "NTA Run Script Error", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return False
            
                
        return True
    
    




    def NTA_Open_Experiment(self, save_directory, sample_name, timeout=2):
        """Opens the experiment given by save_directory and sample_name in the NTA 3.3 program.
        If samples have individual directories the file path is save_directory\\sample_name\\sample_name.nano.
        Otherwise it is save_directory\\sample_name.nano.
        Also checks the video file names loaded in Current Experiment and waits until they
        are similar to sample_name. Will check for up to timeout minutes."""
        
        if not self.NTA_Check_Existence():
            return False
        
        
        nta = self.NTA_app["NTA 3.3"]
        ## Wait for the window to be ready.
        time.sleep(2)
        ## Select the analysis tab.
        nta.TabControl2.Select(u'Analysis')
        time.sleep(2)
        
        ## Check that tab was selected.
        if nta.TabContol2.get_selected_tab() != 2:
            message = "Could not select the Analysis tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Open Experiment Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        ## Select current experiment tab.
        nta.TabControl4.Select(u"Current Experiment")
        time.sleep(2)
        
        ## Check that tab was selected.
        if nta.TabContol4.get_selected_tab() != 1:
            message = "Could not select the Current Experiment tab in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Open Experiment Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        
        ## Before opening the experiment look to see what the experiment name is.
        ## It should be save_directory\sample_name\sample_name.nano or save_directory\sample_name.nano 
        ## but if the date time option is turned on then it will have a time stamp appended to it.
        ## Look in the directory for the file name and see if it has the time stamp or not.
        if self.batch_data.samples_have_individual_directories:
            sample_directory = os.path.join(save_directory, sample_name)
        else:
            sample_directory = save_directory
        files_in_sample_directory = os.listdir(sample_directory)
        sample_file_list = [name for name in files_in_sample_directory if re.match(sample_name+".nano", name)]
        date_stamp_regex = sample_name+" \d\d\d\d-\d\d-\d\d \d\d-\d\d-\d\d.nano"
        date_stamp_list = [name for name in files_in_sample_directory if re.match(date_stamp_regex, name)]
        
        ## If there are multiple nano files in the sample directory with the same name 
        ## choose the one with the greatest timestamp. This also covers the case where 
        ## there are time stamp files and a non stamped file. Choose the highest time stamp 
        ## file in both situations.
        if len(date_stamp_list) > 0:
            file_path = os.path.join(sample_directory, max(date_stamp_list))
            
        elif len(sample_file_list) == 1:
            file_path = os.path.join(sample_directory, sample_file_list[0])
        
        else:
            message = "Could not locate the .nano file for sample " + sample_name + ". Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
            
        
        
        
        ## Click the open experiment button.
        nta["Open Experiment"].Click()
        time.sleep(2)
        
        ## Make sure the click worked and created an Open dialog.
        window_check_result = self.NTA_Window_Check("\'Open\'", True, 0.5)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            message = "Could not open an experiment in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        
        ## File dialog interaction.
        window = self.NTA_app["Open"]
        time.sleep(2)
        window["Edit"].set_text(file_path)
        time.sleep(2)
        
        ## Check that the text was edited.
        if window["Edit"].texts()[0] != file_path:
            message = "Could not enter the file path into the Open dialog in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Open Experiment Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        window = automation.WindowControl(Name = "Open")
        window.SetActive(waitTime=1)
        window.SplitButtonControl(Name = "Open").Click()
        time.sleep(3)
        
        ## Make sure the window closed after clicking open.
        ## If it did not then something went wrong.
        window_check_result = self.NTA_Window_Check("\'Open\'", False, 0.5)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Timed Out":
            message = "Could not open an experiment in NTA. Batch aborted."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        
        
        reg_exp = sample_name + ".*"
        start_time = time.time()
        
        ## Wait for experiment to load in NTA.
        while not self.want_abort:
            
            loaded_samples = nta.ListBox.item_texts()
            if not any([re.match(reg_exp,loaded_sample) for loaded_sample in loaded_samples]):
                time.sleep(1)
                
                if (time.time() - start_time)/60 > timeout:
                    message = "The max wait time of " + str(timeout) + " minutes was reached while waiting for an experiment to load in the NTA 3.3 program. \nCheck the program and try again."
                    msg_dlg = wx.MessageDialog(None, message, "Error. Batch Aborted.", wx.OK | wx.ICON_ERROR)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    return False
                
            else:
                return True
            
        return False
        
        
        





    def NTA_Export_Results(self, timeout = 10):
        """Looks for an Export Settings dialog and clicks Export if it exists 
        otherwise clicks Export Results and then clicks Export on the created dialog.
        
        The Export Settings dialog remains open while export is taking place so the 
        function waits for it to close for up to timeout minutes."""
        
        if not self.NTA_Check_Existence():
            return False
        
        
        nta = self.NTA_app["NTA 3.3"]
        ## Wait for the window to be ready.
        time.sleep(2)
        
        ## Look to see if the script called export or not.
        window_check_result = self.NTA_Window_Check("\'Export Settings\'", True, 0.5)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Success":
            window = automation.WindowControl(Name = "Export Settings")
            window.SetActive(waitTime=1)
            window.ButtonControl(Name = "Export").Click()
            time.sleep(2)
            
        else:
            nta.TabControl2.Select(u'Analysis')
            time.sleep(2)
            
            ## Check that tab was selected.
            if nta.TabContol2.get_selected_tab() != 2:
                message = "Could not select the Analysis tab in NTA. Batch aborted."
                msg_dlg = wx.MessageDialog(None, message, "NTA Export Results Error", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return False
        
        
            ## Select current experiment tab.
            nta.TabControl4.Select(u"Current Experiment")
            time.sleep(2)
            
            ## Check that tab was selected.
            if nta.TabContol4.get_selected_tab() != 1:
                message = "Could not select the Current Experiment tab in NTA. Batch aborted."
                msg_dlg = wx.MessageDialog(None, message, "NTA Export Results Error", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return False
        
        
            ## Click the export results button.
            nta["Export Results"].Click()
            time.sleep(2)
            
            ## Make sure export window appears.
            window_check_result = self.NTA_Window_Check("\'Export Settings\'", True, 0.5)
        
            if window_check_result == "Abort":
                return False
            elif window_check_result == "Time Out":
                message = "Could not export results in NTA. Batch aborted."
                msg_dlg = wx.MessageDialog(None, message, "NTA Export Results Error", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return False
            
            
            window = automation.WindowControl(Name = "Export Settings")
            window.SetActive(waitTime=1)
            window.ButtonControl(Name = "Export").Click()


        
        
        ## Wait for export results window to disappear before proceeding.
        window_check_result = self.NTA_Window_Check("\'Export Settings\'", False, timeout)
        
        if window_check_result == "Abort":
            return False
        elif window_check_result == "Time Out":
            message = "The max wait time of " + str(timeout) + " minutes was reached while waiting for NTA to export results. \nCheck the program and try again. \nBatch Aborted."
            msg_dlg = wx.MessageDialog(None, message, "NTA Export Results Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        elif window_check_result == "Success":
            return True
        
        
        

        return True



    
    
    def CETAC_Check_Existence(self):
        """Checks to make sure the CETAC Workstation program is still open. Returns True if it is, False if not."""
        
        windows = pywinauto.findwindows.find_elements()
        if not any([re.search(r"CETAC Workstation", str(window)) for window in windows]):
            message = "The CETAC Workstation program is not open."
            msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        else:
            return True
        
        
    
    
    




    def CETAC_Close_All_Windows(self):
        """Looks for any windows open under the CETAC Workstation app and closes them."""
        
        if not self.CETAC_Check_Existence():
            return False
        
        ## Close any window controls that are open.
        window_controls = [control for control in self.CETAC_app.GetChildren() if control.ControlTypeName == "WindowControl"]
        for control in window_controls:
            control.SetActive(waitTime=1)
            control.ButtonControl(Name = "Close").Click()
            

        return True

 
    
    
    def CETAC_Run_Script(self):
        """Clicks the Run Script button in the CETAC Workstation program."""
        
        if not self.CETAC_Check_Existence():
            return "CETAC Program CLosed"
        
        self.CETAC_app.SetActive(waitTime=1)
        
        ## Check to see if the button is enabled before clicking.
        if self.CETAC_app.ButtonControl(Name="Run Script").IsEnabled:
            self.CETAC_app.ButtonControl(Name="Run Script").Click()
        else:
            message = "An error occured loading the script. \nThe Run Script button in the CETAC program cannot be clicked. \nBatch Aborted."
            msg_dlg = wx.MessageDialog(None, message, "CETAC Workstation Run Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return "Run Script Button Disabled"
        
        ## Wait a couple of seconds and see if the abort script button becomes enabled.
        ## This is confirmation that the script is running.
        time.sleep(3)
        
        warning_dialogs = [control for control in self.CETAC_app.GetChildren() if control.ControlTypeName == "WindowControl" and control.Name == "Warning"]
        if any([dialog.TextControl().Name == "No tube locations have been selected in \"Select Sample Set\"" for dialog in warning_dialogs]):
            message = "Tube locations have not been selected in the CETAC program. Set tube locations and try again."
            msg_dlg = wx.MessageDialog(None, message, "CETAC Workstation Run Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return "No Samples Selected"
        
        elif self.CETAC_app.ButtonControl(Name="Abort Script").IsEnabled:
            return "Running Script"
        else:
            message = "An unknown error has occured. \nThe Run Script button in the CETAC program has been clicked, but the script is not running."
            msg_dlg = wx.MessageDialog(None, message, "CETAC Workstation Run Script Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return "Abort Script Button Disabled"
    
    
    
    
    
    
    def CETAC_Check_For_Errors(self):
        """This is intended to look for errors generated by the CETAC program, but the error dialogs
         created by the CETAC program are not associated with it, so this function really just looks 
         for dialogs with the title "Error" and returns True if there are any and False otherwise."""
         
        if not self.CETAC_Check_Existence():
            return False
        
        windows = pywinauto.findwindows.find_elements()
        if any([re.search(r"Error", str(window)) for window in windows]):
            message = "The CETAC Workstation program appears to have errored. The batch has been aborted."
            msg_dlg = wx.MessageDialog(None, message, "CETAC Workstation Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return True
        else:
            return False
    

    
 
    
    
    def CETAC_Check_For_COM(self):
        """Checks that communications are extablished with the CETAC Workstation and autosampler by 
        checking to see if the Disconnect button is enabled. Returns True if it is and False otherwise."""
        
        if not self.CETAC_Check_Existence():
            return False
        
        ## Disconnect button only shows up if Home tab is the active tab.
        if len(self.CETAC_app.TabItemControl(Name = "Home").GetChildren()) == 1:
            self.CETAC_app.SetActive(waitTime=1)
            self.CETAC_app.TabItemControl(Name = "Home").Click()
            
            
        if self.CETAC_app.ButtonControl(Name = "Disconnect").IsEnabled:
            return True
        else:
            message = "Communications have not been established with the autosampler in the CETAC program. Initialize communication and try again."
            msg_dlg = wx.MessageDialog(None, message, "CETAC Workstation Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
    
    
    
    
    
    def CETAC_Check_For_Active_Script(self):
        """Checks that a script is opened in the CETAC Workstation program by checking 
        if the Run Script button is enabled. Returns True if it is and False otherwise."""
        
        if not self.CETAC_Check_Existence():
            return False
        
        if not self.CETAC_app.ButtonControl(Name="Run Script").IsEnabled:
            message = "There is no script loaded in the CETAC program. \nLoad a script and try again."
            msg_dlg = wx.MessageDialog(None, message, "CETAC Workstation Error", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return False
        
        else:
            return True
    
    
    
    
    
    
    
    def Listen_For_AS_Signal(self, timeout):
        """Reads the current state of the autosampler output by reading the response from the Arduino.
        Continually sends an \"R\" to the Arduino until a response is received, an abort is 
        detected, a serial exception occurs, or timeout minutes have passed. Returns \"Signal Received\" 
        if the Arduino responds with \"Signal From Autosampler Detected\", \"Abort\" 
        if an abort was detected, \"Pulled Plug\" if a serial exception occured, or \"Time Out\" 
        if timeout minutes have passed without receiving a signal."""
        
        ## Flush input and output buffers.
        try:
            self.ComPort.reset_input_buffer()
        except serial.serialutil.SerialException:
            return "Pulled Plug"
        
        
        ## Listen for a signal from the autosampler until "timeout" minutes have passed.
        start_time = time.time()
        while not self.want_abort:
            try:
                
                self.ComPort.reset_output_buffer()
                self.ComPort.write(b"R")
                
                data = self.ComPort.readline()
                
                
                if len(data) > 0:
                    data = str(data)
                    if re.match(r"b'Signal From Autosampler Detected\\r\\n'", data):
                        return "Signal Received"
        
            
            except serial.serialutil.SerialException:
                return "Pulled Plug"
    
            if (time.time() - start_time)/60 > timeout:
                message = "The max wait time of " + str(timeout) + " minutes was reached while waiting for a signal from the autosampler. \nCheck the Arduino and autosampler script and try again."
                msg_dlg = wx.MessageDialog(None, message, "Error. Batch Aborted.", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return "Time Out"
    
        return "Abort"
    
    
    
    
    
    
    
    def Check_AS_Output_Latch(self, timeout):
        """Reads the latched state of the autosampler output by reading the response from the Arduino.
        Continually sends an \"S\" to the Arduino until a response is received, an abort is 
        detected, a serial exception occurs, or timeout minutes have passed. Returns \"Signal Received\" 
        if the Arduino responds with \"Autosampler Signal Latched As True\", \"Abort\" 
        if an abort was detected, \"Pulled Plug\" if a serial exception occured, or \"Time Out\" 
        if timeout minutes have passed without receiving a signal."""
        
        ## Flush input and output buffers.
        self.ComPort.reset_input_buffer()
        
        
        ## Listen for a signal from the autosampler until "timeout" minutes have passed.
        start_time = time.time()
        while not self.want_abort:
            try:
                
                self.ComPort.reset_output_buffer()
                self.ComPort.write(b"S")
                
                data = self.ComPort.readline()
                
                
                if len(data) > 0:
                    data = str(data)
                    if re.match(r"b'Autosampler Signal Latched As True\\r\\n'", data):
                        return "Signal Received"
        
            
            except serial.serialutil.SerialException:
                return "Pulled Plug"
    
            if (time.time() - start_time)/60 > timeout:
                message = "The max wait time of " + str(timeout) + " minutes was reached while waiting for a signal from the autosampler. \nCheck the Arduino and autosampler script and try again."
                msg_dlg = wx.MessageDialog(None, message, "Error. Batch Aborted.", wx.OK | wx.ICON_ERROR)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return "Time Out"
    
        return "Abort"
    
    
    
    
    
    
    
    
    def Reset_AS_Output_Latch(self):
        """Resets the state of the autosampler output latch in the Arduino.
        Continually sends a \"C\" to the Arduino until a response is received, an abort is 
        detected, a serial exception occurs, or after 10 attempts there is no response from the Arduino. 
        Returns \"Latch Cleared\" if the Arduino responds with \"Autosampler Signal Latch Cleared\", 
        \"Abort\" if an abort was detected, \"Pulled Plug\" if a serial exception occured, or \"Time Out\" 
        if timeout minutes have passed without receiving a signal."""
        
        ## Flush input and output buffers.
        self.ComPort.reset_input_buffer()
        
        
        try:
            
            for i in range(10):
                self.ComPort.reset_output_buffer()
                self.ComPort.write(b"C")
                
                data = self.ComPort.readline()
                
                if len(data) > 0:
                    data = str(data)
                    if re.match(r"b'Autosampler Signal Latch Cleared\\r\\n'", data):
                        return "Latch Cleared"
    
                
                time.sleep(1)
                if i == 9:
                    return "Time Out"
                
                if self.want_abort:
                    return "Abort"
    
        
        except serial.serialutil.SerialException:
            return "Pulled Plug"


    
    
    
    
    
    
    def Send_Signal_To_AS(self):
        """Sends a trigger signal to the autosampler through the Arduino.
        Continually sends a \"T\" to the Arduino until a response is received, an abort is 
        detected, a serial exception occurs, or after 10 attempts there is no response from the Arduino. 
        Returns \"Signal Sent\" if the Arduino responds with \"Signal Sent to Autosampler\", 
        \"Abort\" if an abort was detected, \"Pulled Plug\" if a serial exception occured, or \"Time Out\" 
        if timeout minutes have passed without receiving a signal."""
        
        ## Flush input and output buffers.
        self.ComPort.reset_input_buffer()
        
        ## Look for confirmation signal from Arduino.
        ## Try for 10 seconds to get signal from Arduino before assuming an error.
        try:
            
            for i in range(10):
                self.ComPort.reset_output_buffer()
                self.ComPort.write(b"T")
                
                data = self.ComPort.readline()
                
                if len(data) > 0:
                    data = str(data)
                    if re.match(r"b'Signal Sent To Autosampler\\r\\n'", data):
                        return "Signal Sent"
    
                
                time.sleep(1)
                if i == 9:
                    return "Time Out"
                
                if self.want_abort:
                    return "Abort"
                
        
        except serial.serialutil.SerialException:
            return "Pulled Plug"
    
    
    
    
    
## Commenting out for now. This is not used anywhere and needs to be slightly modified if it ever needs to be used.
## 11-08-2018
#    def Listen_For_Generic_NTA_Signal(self, timeout = 10):
#        """Looks for a dialog created by the NTA 3.3 program with the title \"Script Message\" and 
#        clicks the OK button on it if found. The function keeps looking until timeout minutes 
#        have passed. Returns \"Message Received\" if a message was found or \"Time Out\" if 
#        timeout minutes have passed."""
#        
#        start_time = time.time()
#        
#        ## Find arbitrary message from NTA 3.3
#        while not any([re.search(r"\'Script Message\'", str(window)) for window in self.NTA_app.windows()]):
#            time.sleep(5)
#            
#            if (time.time() - start_time)/60 > timeout:
#                message = "The max wait time of " + str(timeout) + " minutes was reached while waiting for a signal from the NTA 3.3 program. \nCheck the program and try again."
#                msg_dlg = wx.MessageDialog(None, message, "Error. Batch Aborted.", wx.OK | wx.ICON_ERROR)
#                msg_dlg.ShowModal()
#                msg_dlg.Destroy()
#                return "Time Out"
#        
#        ## TODO. This type of signal can't instantly close the message box or the rest of the NTA script will proceed.
#        ##       This won't be used at time of writing but if we ever need this signal this function will need to change.
#        ## Close the message box.
#        self.NTA_app["Script Message"].OK.click()
#        return "Message Received"
        
        
        
        
        
        
    def Listen_For_End_Of_Script_NTA_Signal(self, timeout = 10):
        """Looks for a dialog created by the NTA 3.3 program with the title \"NTA\" and 
        clicks the OK button on it if found. The function keeps looking until timeout minutes 
        have passed or an abort was detected. Returns \"Signal Received\" if a message was found,
        \"Aborted" if an abort was detected, or \"Time Out\" if timeout minutes have passed."""
        
        
        window_check_result = self.NTA_Window_Check("\'NTA\'", True, timeout)
        
        if window_check_result == "Abort":
            return "Aborted"
        elif window_check_result == "Time Out":
            message = "The max wait time of " + str(timeout) + " minutes was reached while waiting for a signal from the NTA 3.3 program. \nCheck the program and try again."
            msg_dlg = wx.MessageDialog(None, message, "Error. Batch Aborted.", wx.OK | wx.ICON_ERROR)
            msg_dlg.ShowModal()
            msg_dlg.Destroy()
            return "Time Out"
        elif window_check_result == "Success":
            ## Click the ok button on the end of script message.
            self.NTA_app["NTA"].OK.click()
            return "Signal Received"
        






    def NTA_Abort_Script(self):
        """Clicks the Abort button on the script panel in the NTA 3.3 program."""
        
        if not self.NTA_Check_Existence():
            return
        
        nta = self.NTA_app["NTA 3.3"]
        ## Wait for the window to be ready.
        time.sleep(2)
        ## Look to see if the script panel is pulled out or not.
        try:
            if any([re.match(r"PanelClass", str(window.class_name())) for window in self.NTA_app.windows()]):
                self.NTA_app.PanelClass.Abort.click()
                
            else:
                ## Click the abort button.
                nta["Abort"].click()
        except Exception as e:
            print("Caught exception when trying to abort NanoSight script.")
            print(str(e))





    def CETAC_Abort_Script(self):
        """Clicks the Abort Script button in the CETAC Workstation program."""
        
        if not self.CETAC_Check_Existence():
            return
        
        self.CETAC_app.SetActive(waitTime=1)
        self.CETAC_app.ButtonControl(Name="Abort Script").Click()







    def NTA_Window_Check(self, regex, find_window, timeout):
        """Searches for an NTA window that matches the regex for timeout minutes if
        find_window is True. If find_window is False then waits for up to timeout minutes
        for the window to disappear.
        Returns Time Out if timedout, Abort if want_abort was set during the search time,
        or Success if a window was found or the window closed depending on find_window."""
        
        start_time = time.time()
        while not self.want_abort:
            try:
                window_found = any([re.search(regex, str(window)) for window in self.NTA_app.windows()])
                
            except Exception:
                time.sleep(1)
                
            else:
                if find_window and not window_found:
                    time.sleep(1)
                elif not find_window and window_found:
                    time.sleep(1)
                else:
                    return "Success"
                
            if (time.time() - start_time)/60 > timeout:
                return "Time Out"


        return "Abort"









class Automation_GUI(wx.Frame):
    
    
    def __init__(self, parent, title, *args, **kwargs):
        
        super(Automation_GUI, self).__init__(parent, title=title, size = (725, 375), *args, **kwargs) 
        
        ## Set up event handlers for serial and batch threads.
        self.Connect(-1, -1, EVT_PULLED_PLUG, self.pulled_plug)
        self.Connect(-1, -1, EVT_THREAD_ABORTED, self.thread_aborted)
        
        self.InitUI()
        self.Centre()
        self.Show()

        
    def InitUI(self):
        
        ## Set some default settings.
        self.batch_thread = None
        
        
        ##############
        ## Menu Bar
        ##############
        menubar = wx.MenuBar()
        fileMenu = wx.Menu()
        
        ## Add items to file menu.
        about_item = fileMenu.Append(100, '&About   \tCtrl+A', 'Information about the program')
        self.Bind(wx.EVT_MENU, self.OnAbout, about_item)
                
        fileMenu.AppendSeparator()
        
        quit_item = wx.MenuItem(fileMenu, 101, 'Quit', 'Quit Application')
        fileMenu.Append(quit_item)
        self.Bind(wx.EVT_MENU, self.OnQuit, quit_item)
        
        ## Add file menu to the menu bar.
        menubar.Append(fileMenu, "&File")
        
        menubar.MacSetCommonMenuBar(menubar)
        ## Put menu bar in frame.
        self.SetMenuBar(menubar)
        
        
        ## Create main panel.
        panel = wx.Panel(self)
        
        ## Create common sizerflags object.
        sizer_flags = wx.SizerFlags(1)
        sizer_flags.Align(wx.ALIGN_LEFT).Border(wx.LEFT | wx.RIGHT | wx.TOP, 10)
        
        main_vbox = wx.BoxSizer(wx.VERTICAL)
        
        ## Add a status bar to the bottom for reporting messages.
        #self.statusbar = self.CreateStatusBar()
        
        ## TODO: Possible additional feature to add program modes to acquire only, process only, or acquire then process (as a batch).
        ## Add radio buttons to select the program mode, acquire or process.
        #mode_st = wx.StaticText(panel, label="Program Mode:")
        
        #########################
        ## Sample List File
        #########################
        ## Add label, button, and feedback for sample list file.
        sample_list_st = wx.StaticText(panel, label="Sample List File (.csv or Excel):")
        main_vbox.Add(sample_list_st, flag = wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.TOP, border = 10)
        
        sample_hbox = wx.BoxSizer(wx.HORIZONTAL)
        
        self.sample_list_tc = wx.TextCtrl(panel, value = "", style = wx.TE_READONLY)
        sample_hbox.Add(self.sample_list_tc, flag = wx.ALIGN_LEFT | wx.EXPAND | wx.LEFT | wx.RIGHT, border = 10, proportion = 1)
        
        self.sample_list_button = wx.Button(panel, label = "Open")
        self.sample_list_button.Bind(wx.EVT_BUTTON, self.On_Sample_List_Open)
        sample_hbox.Add(self.sample_list_button, flag = wx.ALIGN_LEFT)
        
        main_vbox.Add(sample_hbox, flag = wx.ALIGN_LEFT | wx.BOTTOM | wx.EXPAND, border = 10)
        
        
        
        #########################
        ## Individual Directory Checkbox
        #########################
        self.individual_directories_checkbox = wx.CheckBox(panel, label = "Create Individual Directories For Each Sample")
        self.individual_directories_checkbox.SetValue(True)
        self.samples_have_individual_directories = True
        
        self.individual_directories_checkbox.Bind(wx.EVT_CHECKBOX, self.OnCheck)
        
        main_vbox.Add(self.individual_directories_checkbox, flag = wx.ALIGN_LEFT | wx.LEFT | wx.BOTTOM, border = 10)
        
        
        
        ##########################
        ## Number of Samples
        ##########################
        ## Add a label to show the number of samples.
        num_samples_hbox = wx.BoxSizer(wx.HORIZONTAL)
        
        num_of_samples_label_st = wx.StaticText(panel, label="Number of Samples:")
        num_samples_hbox.Add(num_of_samples_label_st, flag = wx.ALIGN_LEFT | wx.ALL, border = 10)
        
        self.num_of_samples_st = wx.StaticText(panel, label="")
        num_samples_hbox.Add(self.num_of_samples_st, flag = wx.ALIGN_LEFT | wx.ALL, border = 10)
        
        main_vbox.Add(num_samples_hbox, flag = wx.ALIGN_LEFT)
        
        
        
        ##########################
        ## Sample Progress
        ##########################
        ## Add text box to display what sample is in progress.
        run_progress_st = wx.StaticText(panel, label="Run Progress:")
        main_vbox.Add(run_progress_st, flag = wx.ALIGN_LEFT | wx.LEFT | wx.RIGHT | wx.TOP, border = 10)
        
        self.list_ctrl_col_names = ["Sample Name", "Save Directory", "Acquire Script", "Process Script",
                                    "Acquisition Progress", "Processing Progress"]
        
        self.sample_list_ctrl = wx.ListCtrl(panel, style=wx.LC_REPORT)
        
        
        for col_name in self.list_ctrl_col_names:    
            self.sample_list_ctrl.InsertColumn(self.list_ctrl_col_names.index(col_name), col_name, width = (len(col_name)*5 + 35))
            
        main_vbox.Add(self.sample_list_ctrl, flag = wx.ALIGN_LEFT | wx.EXPAND | wx.LEFT | wx.RIGHT | wx.BOTTOM, border = 10)
        
        
        ##########################
        ## Start Batch Button
        ##########################
        ## Add button to toggle starting and stopping run.
        self.toggle_button = wx.ToggleButton(panel, label = "Start Batch")
        self.toggle_button.Disable()
        self.toggle_button.Bind(wx.EVT_TOGGLEBUTTON, self.On_Toggle)
        
        
        main_vbox.Add(self.toggle_button, flag = wx.ALIGN_CENTER | wx.ALL, border = 10)
        

        panel.SetSizer(main_vbox)
        panel.Fit()
        self.Fit()
        




    def OnQuit(self, e):
        self.Close()



    def OnAbout(self, event):
        message = "Author: Travis Thompson \n\n" + \
        "Creation Date: November 2018" + \
        "Version Number: " + __version__
       
        dlg = GMD.GenericMessageDialog(None, message, "About", wx.OK | wx.ICON_INFORMATION)
        dlg.ShowModal()
        dlg.Destroy()





    def On_Sample_List_Open(self, event):
        """Creates the open file dialog box and gets a file from the user. The file 
        is then checked to make sure it is formatted correctly. The file is checked 
        to make sure it has the correct columns, it is the correct file type, and 
        that it has values in every column. Also makes sure that all file paths in 
        the file exist. If all checks pass then the GUI is updated with the 
        appropriate information."""
        
        message = "Select Sample List File (.csv or Excel)"
        dlg = wx.FileDialog(None, message = message, style=wx.FD_OPEN | wx.FD_CHANGE_DIR)
    
        if dlg.ShowModal() == wx.ID_OK:
            filepath = dlg.GetPath()
            dlg.Destroy()
            
            if not re.match(r".*\.xlsx|.*\.xlsm|.*\.xls|.*\.csv", filepath):
                message = "Incorrect file type chosen. Please select an Excel or .csv file."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
                return
                
            else:
                if re.match(r".*\.csv", filepath):
                    try:
                        sample_df = pandas.read_csv(filepath)
                    except Exception:
                        message = "The selected Sample List File could not be opened. This could be due to the file being corrupted or open in another program."
                        msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                        msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        return
                else:
                    try:
                        sample_df = pandas.read_excel(filepath)
                    except Exception:
                        message = "The selected Sample List File could not be opened. This could be due to the file being corrupted or open in another program."
                        msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                        msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        return
                
                self.toggle_button.Disable()
                    
                required_columns = set(self.list_ctrl_col_names) - set(["Acquisition Progress", "Processing Progress"])
                
                if not set(required_columns).issubset(sample_df.columns) :
                    missing_columns = set(required_columns) - set(sample_df.columns)
                    message = "The selected file is missing columns for: \n\n" + "\n".join(missing_columns)
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    
                elif sample_df.isnull().values.any():
                    message = "The selected file has missing values."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    
                elif not sample_df.loc[:, "Sample Name"].is_unique:
                    message = "The selected file's sample names are not unique."
                    msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                    msg_dlg.ShowModal()
                    msg_dlg.Destroy()
                    
                else:
                    
                    
                    ## Make sure all script files actually exists.
                    file_paths_OK = True
                    
                    acqs_paths = {str(sample):os.path.exists(str(path)) for (sample, path) in sample_df.loc[:, ["Sample Name", "Acquire Script"]].values}
                    
                    if not all(acqs_paths.values()):
                        message = "The Acquire Script file path given for sample(s): \n\n" + "\n".join([sample for sample, value in acqs_paths.items() if not value]) + "\n\nare not valid file path(s). Check that the path(s) exist and try again."
                        msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
                        msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        file_paths_OK = False
                    
    
                    ps_paths = {str(sample):os.path.exists(str(path)) for (sample, path) in sample_df.loc[:, ["Sample Name", "Process Script"]].values}
    
    
                    if not all(ps_paths.values()):
                        message = "The Process Script file path given for sample(s): \n\n" + "\n".join([sample for sample, value in ps_paths.items() if not value]) + "\n\nare not valid file path(s). Check that the path(s) exist and try again."
                        msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
                        msg_dlg.ShowModal()
                        msg_dlg.Destroy()
                        file_paths_OK = False
    
                    
                    
                    
                    
                    
                    if file_paths_OK:
                                            
                        self.sample_list_tc.SetValue(filepath)
                        
                        self.num_of_samples_st.SetLabel(str(len(sample_df)))
                        
                        self.sample_list_ctrl.DeleteAllItems()
                        
                        for i in range(len(sample_df)):
                            self.sample_list_ctrl.InsertItem(i, str(sample_df[self.list_ctrl_col_names[0]][i]))
                            self.sample_list_ctrl.SetItem(i, self.list_ctrl_col_names.index("Acquisition Progress"), "Not Started")
                            self.sample_list_ctrl.SetItem(i, self.list_ctrl_col_names.index("Processing Progress"), "Not Started")
                            
                            for j in range(1,len(required_columns)):
                                self.sample_list_ctrl.SetItem(i, j, str(sample_df[self.list_ctrl_col_names[j]][i]))
                            
                        
                        self.sample_df = sample_df
                        self.toggle_button.Enable()
    
            
        
    
    def OnCheck(self, event):
        """Change the state of the internal variable to match the state of the check box."""
        
#        """If the user clicks the check after opening the Sample List File then the directories 
#        won't get made, so see if the sample list exists and create the directories if the box is checked."""
        
        check_box = event.GetEventObject()
        self.samples_have_individual_directories = check_box.GetValue()
        
#        if check_box.GetValue() and self.toggle_button.Enabled():
#            
#            sample_df = self.sample_df
#            
#            for i in range(len(sample_df)):
#                directory = sample_df.loc[i, "Save Directory"]
#                sample_name = sample_df.loc[i, "Sample Name"]
#                directory = os.path.join(directory, sample_name)
#                    
#                try:
#                    os.makedirs(directory)
#                except OSError as e:
#                    if e.errno != errno.EEXIST:
#                        message = "An error occurred while creating the directory for sample " + sample_name + "."
#                        message = message + "\n\n" + str(e)
#                        msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
#                        msg_dlg.ShowModal()
#                        msg_dlg.Destroy()
#                        return
            
    
    
    
        
    
    
    
    def On_Toggle(self, event):
        """After clicking the "Start/Abort Batch button this function either 
        starts or aborts the batch appropriately."""
        
        button = event.GetEventObject()
        
        
        if button.GetValue() == True:
            ## Disable the button to open a different file so the user can't change the sample list mid operation.
            self.sample_list_button.Disable()
            self.individual_directories_checkbox.Disable()
            
            if not self.batch_thread:
                self.batch_thread = BatchThread(self)
                button.SetLabel("Abort Batch")
            
            else:
                message = "Batch thread already started, something seriously wrong happened."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
            
        
        else:
            
            if self.batch_thread:
                self.batch_thread.abort()
                ## Disable button until thread has been aborted.
                button.Disable()
            
            else:
                message = "Batch thread already aborted, something seriously wrong happened."
                msg_dlg = wx.MessageDialog(None, message, "Warning", wx.OK | wx.ICON_EXCLAMATION)
                msg_dlg.ShowModal()
                msg_dlg.Destroy()
            
            
    






    def pulled_plug(self, event):
        """Handler for PulledPlug events. Creates a message to indicate that the 
        Arduino was disconnected and re-enables all buttons."""
        
        message = "Arduino Disconnected. Batch Aborted."
        msg_dlg = wx.MessageDialog(None, message, "Error", wx.OK | wx.ICON_ERROR)
        msg_dlg.ShowModal()
        msg_dlg.Destroy()
        
        self.toggle_button.SetValue(False)
        self.toggle_button.SetLabel("Start Batch")
        
        self.batch_thread = None
        
        self.sample_list_button.Enable()
        self.individual_directories_checkbox.Enable()
   


    def thread_aborted(self, event):
        """Handler for aborted thread. Re-enables all buttons."""
        
        self.toggle_button.SetValue(False)
        self.toggle_button.SetLabel("Start Batch")
        self.toggle_button.Enable()
        
        self.batch_thread = None

        self.sample_list_button.Enable()
        self.individual_directories_checkbox.Enable()






        
        
def main():

    ex = wx.App(False)
    Automation_GUI(None, title = "NanoSight Automation")
    ex.MainLoop()    


if __name__ == '__main__':
    main()