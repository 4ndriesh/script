import os
import yaml
from logging_err import *

class Config:
    __instance = None

    @staticmethod
    def inst():
        if Config.__instance == None:
            Config.__instance = Config()
        return Config.__instance

    def __init__(self):
        print('config')
        import sys, os
        if getattr(sys, 'frozen', False):
            # If the application is run as a bundle, the pyInstaller bootloader
            # extends the sys module by a flag frozen=True and sets the app
            # path into variable _MEIPASS'.
            self.BASE_DIR =os.getcwd()
        else:
            self.BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


        self.filename_config = os.path.join(self.BASE_DIR, 'setting.yaml')
        try:
            with open(self.filename_config, 'r') as yaml_config_file:
                self.config = yaml.load(yaml_config_file)
        except Exception as e:
            exeption_print(e)

    @property
    def version(self):
        try:
            return [self.config['usecols'],self.BASE_DIR, self.config['path'], self.config['sheet_1'],
                    self.config['sheet_2'], self.config['row_1'], self.config['row_2']]
        except:
            self.config = {'version':'нет версии'}

    @version.setter
    def version(self, value):
        self.config['version'] = value
        with open(self.filename_config, "w") as f:
            yaml.dump(self.config, f,default_flow_style=False)



