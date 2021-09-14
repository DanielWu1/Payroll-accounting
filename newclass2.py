import pandas as pd
import datetime
import xlrd
from newclass import collect as nc

class combine:

    needlist = ['基本工资','职位工资','应出勤','节假日','实出勤','调休','结转下月调休天数','周六加班费',
    '值班','休息出勤','延长时间加班','值班核算','工龄工资','保险补贴','其它','工资核算','核算天数','应发工资','社保','医保',
    '公积金','个税','社保其它','实发工资','单位','部门']

    def __init__(self,gong1,gong2,gong3):
        self.gong1 = gong1
        self.gong2 = gong2
        self.gong3 = gong3
        self.firstcombine = {}
        self.secondcombine= {}
        self.finalcombine = {}
        self.output_sh=''
        self.theTime=''

    def run(self):
        self.combinefirst()
        self.combinesecond()
        #combinethird()
        self.calculate()
        self.write_output()


    def combinefirst(self):
        person_name = ''
        for fname in self.gong1.keys():
            for sname in self.gong2.keys():
                if fname == sname :
                    person_name = fname
                    #print (person_name)
                    self.firstcombine[person_name]=dict(self.gong1[person_name],** self.gong2[person_name])
        #print(self.firstcombine)
        return self.firstcombine

    def combinesecond(self):
        person_name = ''
        for fname in self.firstcombine.keys():
            for sname in self.gong3.keys():
                if fname == sname :
                    person_name = fname
                    #print (person_name)
                    self.secondcombine[person_name]=dict(self.firstcombine[person_name],** self.gong3[person_name])
        #print(self.secondcombine)
        return self.secondcombine

    def calculate (self):
        hesuan = {'工资核算': 0}
        zhiban = {'值班核算': 0}
        hetian = {'核算天数': 0}
        yingfa = {'应发工资': 0}
        shifa = {'实发工资': 0}
        for name in self.secondcombine.keys():
            calhesuan = self.secondcombine[name]['基本工资'] + self.secondcombine[name]['职位工资'] + self.secondcombine[name]['周六加班费']
            #print(hesuan)
            hesuan['工资核算']=calhesuan
            self.secondcombine[name] = dict(self.secondcombine[name],** hesuan)

        #print (self.secondcombine)
        for name in self.secondcombine.keys():
            calzhiban = self.secondcombine[name]['基本工资'] / self.secondcombine[name]['应出勤'] * self.secondcombine[name]['值班'] * 2
            #print(hesuan)
            zhiban['值班核算']=calzhiban
            self.secondcombine[name] = dict(self.secondcombine[name],** zhiban)

        #print (self.secondcombine)

        for name in self.secondcombine.keys():
            calhetian = self.secondcombine[name]['实出勤'] + self.secondcombine[name]['调休'] + self. secondcombine[name]['节假日'] - self.secondcombine[name]['休息出勤'] - self.secondcombine[name]['值班']
            #print(hesuan)
            hetian['核算天数']=calhetian
            self.secondcombine[name] = dict(self.secondcombine[name],** hetian)

        #print(self.secondcombine)
        for name in self.secondcombine.keys():
            calyingfa = self.secondcombine[name]['工资核算'] / self.secondcombine[name]['应出勤'] * self.secondcombine[name]['核算天数'] + self.secondcombine[name]['值班核算'] + self.secondcombine[name]['其它']
            #print(hesuan)
            yingfa['应发工资']=calyingfa
            self.secondcombine[name] = dict(self.secondcombine[name],** yingfa)

        #print (self.secondcombine)

        for name in self.secondcombine.keys():
            calshifa = self.secondcombine[name]['应发工资'] + self.secondcombine[name]['工龄工资'] + self.secondcombine[name]['保险补贴'] - self.secondcombine[name]['社保'] - self.secondcombine[name]['医保'] - self.secondcombine[name]['公积金'] - self.secondcombine[name]['个税'] - self.secondcombine[name]['社保其它']
            #print(hesuan)
            shifa['实发工资']=calshifa
            self.secondcombine[name] = dict(self.secondcombine[name],** shifa)

        #print (self.secondcombine)
        total = {}
        totalin = {}
        for i in self.needlist:
            totalin[i] = 0
        total['合计'] = totalin
        #print (total)
        for name in self.secondcombine.keys():
            for index in self.secondcombine[name]:
                if index == '岗位': next
                else : total['合计'][index] += self.secondcombine[name][index] 

        #print (total)
        #print (self.secondcombine)
        self.secondcombine = dict(self.secondcombine,** total)
        print(self.secondcombine)

    '''def get_time(self):
        ISOTIMEFORMAT = '%Y-%m-%d'
        self.theTime = datetime.datetime.now().strftime(ISOTIMEFORMAT)
        print (self.theTime)
        #return theTime'''



    def write_output(self):
        '''
        indexc = 1
        indexr = 5
        for i in self.needlist:
            self.output_sh.sell
        for name in self.secondcombine.keys():
        '''

        namelist = []
        for name in self.secondcombine.keys():
            namelist.append(name)
            #namelist.append("日期")
        #namelist.append('合计')
        #print (namelist)
        #timelist = [self.get_time]
        #print(timelist)

        
        finaldata = pd.DataFrame(data = self.secondcombine.values(), index = namelist , columns = self.needlist)
        #timedata = pd.DataFrame(data = timelist,columns = None )
        #print (timedata)
        #finaldata = finaldata.set_value(len(finaldata),self.get_time)
        #finaldata = finaldata.T
        #print (finaldata)
        #timedata = pd.DataFrame(data = self.get_time)
        finaldata.to_excel('新建工资统计表.xls' , startrow = 1 ,sheet_name='工资')
        



        #self.zonggong = dict(self.gong1,** self.gong2)
        #print(self.zonggong)

    #def runit(self):
    #    combinefirst (gong1,gong2)


if __name__=='__main__':
    x = nc('2021年3月工资考勤【生产】4.7(1).xlsx', "结构")
    x.run_jiegou()
    #x.get_jiegou()
    #print(x.all_people_map_one)
    y = nc('2021年3月工资考勤【生产】4.7(1).xlsx', "考勤")
    y.run_kaoqin()
    #y.get_kaoqin()
    z = nc('2021年3月工资考勤【生产】4.7(1).xlsx', "补贴")
    z.run_butie()
    #z.get_butie()
    a = combine(x.all_people_map_one,y.all_people_map_two,z.all_people_map_three)
    a.run()
    #a.get_time()