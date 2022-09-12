from openpyxl import load_workbook
from Fetcher import FunctionFetcher

class GetFunctions:

    def __init__(self):
        self.captain_names = ["COMANDANTE/OIM", "OIM/COMANDANTE", "OIM COMANDANTE", "COMANDANTE OIM"]
        self.choff_names = ["IMEDIATO", "CHIEF OFFICER"]
        self.dpos_names = ["OPERADOR DE POSIC DINAMICO", "OPERADORA DE POSIC DINAMICO"]
        self.sdpos_names = ["OPERADOR DE POSIC DINAMICO SR", "OPERADORA DE POSIC DINAMICO SR"]
        self.mnc_names = ["MARINHEIRO DE CONVÉS", "MARINHEIRO DE CONVES"]
        self.rigsup_names = ["SUPERINTENDENTE DE PLATAFORMA"]
        self.maintcoord_names = ["COORD MANUTENÇÃO", "COORDENADOR DE MANUTENÇÃO", "MAINTENANCE COORDINATOR", "COORD DE MANUTENÇÃO"]
        self.elecsup_names = ["SUPERVISOR DE ELETROELETRONICA", "ELECTRICAL SUPERVISOR"]
        self.rop_names = ["RADIO OPERADOR", "RADIO OPERADORA"]
        self.craneop_names = ["OPERADOR DE GUINDASTE", "OP. DE GUINDASTE"]
        self.oqm_names = ["SEGUNDO OFICIAL DE MÁQUINAS", "SEGUNDO OFICIAL DE MAQUINAS"]
        self.cheng_names = ["CHEFE DE MÁQUINAS", "CHEFE DE MAQUINAS", "CHIEF ENGINEER"]
        self.auxplat_names = ["AUXILIAR DE PLATAFORMA", "AUX. PLATAFORMA"]

    def get_data(self, workbook_path):
        data = {}
        self.ff = FunctionFetcher(workbook_path)
        data["CPT"] = self.ff.get_duty(self.captain_names)
        data["IMT"] = self.ff.get_duty(self.choff_names)
        data["DPOS"] = self.ff.get_duty(self.dpos_names)
        data["SDPOS"] = self.ff.get_duty(self.sdpos_names)
        data["MNCS"] = self.ff.get_duty(self.mnc_names)
        data["RIGSUP"] = self.ff.get_duty(self.rigsup_names)
        data["MAINTCOORD"] = self.ff.get_duty(self.maintcoord_names)
        data["ELECSUP"] = self.ff.get_duty(self.elecsup_names)
        data["ROPS"] = self.ff.get_duty(self.rop_names)
        data["CRANEOPS"] = self.ff.get_duty(self.craneop_names)
        data["OQMS"] = self.ff.get_duty(self.oqm_names)
        data["CFM"] = self.ff.get_duty(self.cheng_names)
        data["AUXPLATS"] = self.ff.get_duty(self.auxplat_names)
        data["RM"] = "Victor Borba"
        return data

