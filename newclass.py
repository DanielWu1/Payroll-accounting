import pandas as pd


class collect:

    need_list = ["岗位", "上月", "基本工资", "职位工资", "周六加班费", "姓名", "应出勤", "调休", "值班",'节假日',
    "休息出勤","延长时间加班","实出勤","结转下月调休天数","岗位", "工龄工资", "保险补贴", "其它","社保","医保","公积金","个税","社保其它"]
    #map_need = {}
    #all_people_map={}

    def __init__(self, file_path, sheetname):
        self.df = pd.read_excel(file_path,sheet_name=sheetname)
        #self.df = file_path
        self.df = self.df.T
        self.all_people_map_one={}
        self.all_people_map_two={}
        self.all_people_map_three={}
        #self.map_need={}
        self.flag = False
        #self.nan_value_flag = False

    def delspa (self):
        for name in self.df.columns:
            for element in self.df[name]:
                if(element == "序号"): self.flag = not self.flag
                if (element == "序"): self.flag = not self.flag

            if(not self.flag): del self.df[name]
        #print(self.df)

    def findtop (self):
        map_need={}
        for index in range(len(self.df[self.df.columns[0]])):
            top_value = self.df[self.df.columns[0]][index]
            #if show_chart: print(top_value)
            if top_value in self.need_list:
                map_need[index] = top_value
        self.map_need=map_need
    
    def savedata_jiegou(self):
        for name in self.df.columns:
            person_name = ""
            person_info = {}

            for index in range(len(self.df[name])):
                #if(self.nan_value_flag): break
                if(index in self.map_need):

                    if(self.map_need[index] == "上月"): person_name = self.df[name][index]
                    else: person_info[self.map_need[index]] = self.df[name][index]

            #if show_chart: 
                #print(self.df[name][index], end = " ")
            if person_name == "上月": next
            else: self.all_people_map_one[person_name] = person_info

    def savedata_kaoqin(self):
        for name in self.df.columns:
            person_name = ""
            person_info = {}

            for index in range(len(self.df[name])):
                #if(nan_value_flag): break
                if(index in self.map_need):

                    if(self.map_need[index] == "姓名"): person_name = self.df[name][index]
                    else: person_info[self.map_need[index]] = self.df[name][index]

            #if show_chart: 
                #print(self.df[name][index], end = " ")
            if person_name =='姓名': next
            else :self.all_people_map_two[person_name] = person_info



    def savedata_butie(self):
        for name in self.df.columns:
            person_name = ""
            person_info = {}

            for index in range(len(self.df[name])):
                #if(nan_value_flag): break
                if(index in self.map_need):

                    if(self.map_need[index] == "姓名"): person_name = self.df[name][index]
                    else: person_info[self.map_need[index]] = self.df[name][index]

            #if show_chart: 
                #print(self.df[name][index], end = " ")
            if person_name == "姓名": next
            else: self.all_people_map_three[person_name] = person_info


    def run_jiegou(self):
        self.delspa()
        self.findtop()
        self.savedata_jiegou()
        return(self.all_people_map_one)

    def run_kaoqin(self):
        self.delspa()
        self.findtop()
        self.savedata_kaoqin()
        return(self.all_people_map_two)

    def run_butie(self):
        self.delspa()
        self.findtop()
        self.savedata_butie()
        return(self.all_people_map_three)

    #def get_df(self):
    #    return(self.df)
    def get_jiegou(self):
        #print(self.all_people_map_one)
        return(self.all_people_map_one)

    def get_kaoqin(self):
        #print(self.all_people_map_two)
        return(self.all_people_map_two)

    def get_butie(self):
        #print(self.all_people_map_three)
        return(self.all_people_map_three)


    def test(self):
        print(self.need_list)
        print(self.map_need)
                


if __name__ == '__main__':
    x = collect('工作簿1.xls', "结构")
    x.run_jiegou()
    #x.get_df()
    #print('/n')
    ##x.get_jiegou()
    #x.test()
    y = collect('工作簿1.xls', "考勤")
    y.run_kaoqin()
    ##y.get_kaoqin()
    z = collect('工作簿1.xls', "补贴")
    z.run_butie()
    ##z.get_butie()



