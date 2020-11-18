#
from openpyxl import Workbook
from openpyxl import load_workbook
import re
from pyecharts.charts import Bar,Line,Pie
from pyecharts import options as opts

# names=['李丰祥','尉文革','殷超','李燕','张小春','单军','陈强','杨伟政','马忠英','李忠利','王玉宝','王飞','王龙','集客','未知']         #统计图表
names=['李丰祥','尉文革','殷超','李燕','张小春','单军','陈强','杨伟政','马忠英','李忠利','王玉宝','王飞','张小春','王龙','集客','未知']    #统计日报
def guzhang(date):
    print("故障")
    workbook=load_workbook(filename=r'C:\Users\Administrator\Desktop\家宽末次回复报表明细(新)_11.'+str(date)+'.xlsx')
    sheet1=workbook.active
    # sheet1.title="条形统计图"
    num=""
    for nums in sheet1["V"]:
        num+=nums.value
    # print(num)
    number1=[]
    for n in range(0,len(names)):
        number1.append(len(re.findall(names[n],num)))
    print(number1)
    return number1
def ruoguang(date):
    print("弱光")
    workbook = load_workbook(filename=r'C:\Users\Administrator\Desktop\智能网关强弱光明细_11.'+str(date)+'.xlsx')
    sheet1 = workbook.active
    num = ""
    for nums in sheet1["N"]:
        num += nums.value
    # print(num)
    number2 = []
    for n in range(0, len(names)):
        number2.append(len(re.findall(names[n], num)))
    print(number2)
    return number2
def zhuangji(date):
    print("装机")
    workbook = load_workbook(filename=r'C:\Users\Administrator\Desktop\家宽装机全量指标明细（新）_11.'+str(date)+'.xlsx')
    sheet1 = workbook.active
    num = ""
    for nums in sheet1["R"]:
        num += nums.value
    # print(num)
    number3 = []
    for n in range(0, len(names)):
        number3.append(len(re.findall(names[n], num)))
    print(number3)
    return number3
def zhijian(date):
    print("质检")
    workbook = load_workbook(filename=r'C:\Users\Administrator\Desktop\家宽装机质检明细（新）_11.'+str(date)+'.xlsx')
    sheet1 = workbook.active
    num = ""
    for nums in sheet1["M"]:
        num += nums.value
    # print(num)
    number4 = []
    for n in range(0, len(names)):
        number4.append(len(re.findall(names[n], num)))
    print(number4)
    return number4

def ribao(date,number1,number2,number3,number4):
    worbook = load_workbook(filename=r'C:\Users\Administrator\Desktop\日通报.xlsx')
    sheet = worbook.active
    sheet["O2"].value=date
    for i in range(0, len(number1)):
        sheet["M" + str(i + 4)].value = number1[i]          #故障
        sheet["G" + str(i + 4)].value = number2[i]          #弱光
        sheet["K" + str(i + 4)].value = number3[i]          #装机
        sheet["I" + str(i + 4)].value = number4[i]          #质检
    zimu=["G","H","I","J","K","L","M","N"]
    for zim in zimu:
        sheet[zim+"16"].value=""
    worbook.save(filename=r'C:\Users\Administrator\Desktop\日通报.xlsx')
    print("文件保存成功")

#柱状图
def fengxiB(date,number1,number2,number3,number4):
    bar=Bar()
    bar.add_xaxis(names)
    bar.add_yaxis("故障",number1)
    bar.add_yaxis("弱光",number2)
    bar.add_yaxis("装机",number3)
    bar.add_yaxis("质检",number4)
    bar.set_global_opts(title_opts=opts.TitleOpts(title=str(date)+"号日报分析"),xaxis_opts=opts.AxisOpts(axislabel_opts={"interval":"0"}))
    bar.render("日报分析.html")
#折线图
def fengxiL(date,number1,number2,number3,number4):
    line=Line()
    line.add_xaxis(names)
    line.add_yaxis("故障", number1)
    line.add_yaxis("弱光", number2)
    line.add_yaxis("装机", number3)
    line.add_yaxis("质检", number4)
    line.set_global_opts(title_opts=opts.TitleOpts(title=str(date) + "号日报分析"),xaxis_opts=opts.AxisOpts(axislabel_opts={"interval": "0"}))
    line.render("日报分析.html")
#饼图
# pie=Pie()
# pie.add("", [list(name) for name in zip((names),(number))])
# # pie.set_global_opts(title_opts=opts.LegendOpts(pos_top=20))
# pie.set_series_opts(label_opts=opts.LabelOpts(formatter="{b}:{c}"))
# pie.render("故障分析.html")

def main():
    date=input("请输入查询日期xx日：")
    number1=guzhang(date)
    number2=ruoguang(date)
    number3=zhuangji(date)
    number4=zhijian(date)
    # fengxiB(date,number1,number2,number3,number4)
    # fengxiL(date, number1, number2, number3, number4)
    ribao(date,number1, number2,number3,number4)



if __name__=="__main__":
    main()










