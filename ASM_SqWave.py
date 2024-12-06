import subprocess
import os
from time import sleep

class sqWave:

    def __init__(self,frequency):

        self.frequency=frequency
        
        # Convert frequency to delay and then to hex
        delay_value = hex(int((1 / frequency) * 1000000)) 


        with open('[path to file]') as file:
            lines=[line.rstrip() for line in file.readlines()]

        for line in lines:
            if 'variableName' in line:
                lines[lines.index(line)]='variableName: .asciz' +newDel
        
        with open('/home/user/Desktop/Embedded/Upper/upper.s','w') as file:
            for line in lines:
                file.writelines(line)
                file.write('\n')

    def start(self):

        os.chdir() #to navigate to the directory
        process=subprocess.Popen('make',shell=False,stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
        process.wait()
        self.p1 = subprocess.Popen('./BlinkLed', shell=False, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)

    def stop(self):
        self.p1.terminate()

    def setFrequency(self, frequency):
        # Update the frequency and restart the square wave
        self.__init__(frequency)
        self.start()
