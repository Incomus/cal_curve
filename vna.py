import pyvisa
import time
import re


class VNAController:

    def __init__(self, gpib_address):
        self.rm = pyvisa.ResourceManager()
        self.vna = None
        self.gpib_address = gpib_address
        self.resource_string = f"GPIB0::{gpib_address}::INSTR"
        self.screen_output=None
        self.unit ="GHZ"
        self.logm_or_phase = "PHASE"
        self.connect()


    def find_and_connect(self):
        print("The address isn't 13, try to find and connect the relevant VNA")
        print(self.rm.list_resources())
        resources = [r for r in self.rm.list_resources() if "GPIB" in r]
        print(f"The resources are: {resources}")
        expected_vna_id = '8722D'
        for resource in resources:
            try:
                self.vna = self.rm.open_resource(resource)
                time.sleep(0.5)
                idn = self.vna.query("IDN?")
                if expected_vna_id in idn:
                    print(f"Connected to: {idn}")
                    return True
                else:
                    self.vna.close()
            except pyvisa.VisaIOError as ex:
                print(f"Failed to communicate with {resource}: {ex}")
                return False
        return False


    def connect(self):
        try:
            self.vna = self.rm.open_resource(self.resource_string)
            idn = self.vna.query("IDN?")
            print(f"Connected to: {idn}")
            return True
        except:
            return self.find_and_connect()

    def set_ampl_or_phase(self,logm_or_phase):
        self.logm_or_phase = logm_or_phase
        if logm_or_phase == "PHASE":
            logm_or_phase = "PHAS"
        self.vna.write(f"{logm_or_phase}")

    def set_channel(self,channel_num):

        self.vna.write(f"{channel_num}")

    def set_frequency(self, start_freq, stop_freq,unit):
        self.vna.write(f"STAR {start_freq}{unit}")
        self.vna.write(f"STOP {stop_freq}{unit}")

        self.unit = unit
    def set_span(self, span_val):
        self.vna.write(f"SPAN {span_val}HZ")

    def set_center(self, center_val):
        self.vna.write(f"CENT {center_val}GHZ")

    def get_phase(self):
        """
        Get phase data of the current trace
        """
        pattern = r'LBphase.*?;LB(.*?) '
        phase_values = re.findall(pattern, self.screen_output)[0]
        pattern = 'LBREF (.*?) '
        ref_values = re.findall(pattern, self.screen_output)[0]
        print('Phase Values:', phase_values)
        print('REF Values:', ref_values)
        return str(phase_values),str(ref_values)

    def set_data_to_memory(self):
        self.vna.write("DATI")  # data to memory data -> memory
        self.vna.write("DISPDDM")  # Data divided by memory (linear division, log subtraction).

    def set_number_of_points(self,number_of_points):

        number_of_points = number_of_points.strip()  # Removes leading and trailing whitespaces
        num_points = int(number_of_points)  # Converts the cleaned string to an integer
        self.vna.write(f"POIN {int(number_of_points)}")

    def get_screen_output(self):
        self.vna.write("FORM4")
        self.vna.write("OUTPPLOT")
        time.sleep(1)
        self.screen_output = self.vna.read()

    def put_marker(self,marker_number,at_freq):
        at_freq = at_freq * 1e9
        time.sleep(0.09)
        self.vna.write(f"MARK{marker_number} {at_freq}")
        self.last_active_marker = marker_number

    def get_mark_value(self, mark_n):

        self.vna.write(f"MARK{mark_n}")
        time.sleep(0.1)
        self.vna.write("OUTPMARK")
        data_markers = self.vna.read()
        data_markers_values = data_markers.strip().split(',')
        marker_value = round(float(data_markers_values[0]),2)

        # frequency_value =float(data_markers_values[2])
        # if self.unit == "GHZ":
        #     frequency_value = frequency_value / 1e9
        # elif self.unit =="MHZ":
        #     frequency_value = frequency_value / 1e6
        # frequency_value = round(frequency_value,4)
        return marker_value