from openpyxl import load_workbook


class FunctionFetcher:

    def __init__(self, workbook_path):
        self.pob = load_workbook(workbook_path)

        self.pob_list = self.pob["POB N08"]
        self.function_list = self.pob_list["C5:D200"]


    def get_duty(self, duty_list: list[str]):
        result = []
        for function in self.function_list:
            for name in duty_list:
                if function[1].value:
                    if function[1].value.strip() == name:
                        result.append(self.split_name(self.pob_list["B" + str(function[1].row)].value))
                elif function[0].value: 
                    if function[0].value.strip() == name:
                        result.append(self.split_name(self.pob_list["B" + str(function[0].row)].value))
        return result

    def split_name(self, name):
        if len(name.split()[1]) <= 2:
            return name.split()[0] + " " + name.split()[1] + " " + name.split()[2] 
        else:
            return name.split()[0] + " " + name.split()[1]

