# -*-coding:utf-8-*-
__author__='wanghaibo'
__Mail__='13901751671@139.com'
__version__='0.03'
__date__ ="20170311"

import requests
import logging
import datetime
logging.getLogger('opm reporter').setLevel(logging.DEBUG)
import xlsxwriter
import yaml

def _get_value(obj, key, default='0'):
    str_obj = obj.get(key, default)
    if isinstance(str_obj, str): return unicode(str_obj.decode().encode('utf-8'))
    return str_obj


def _get_one_performance_monitor(DISPLAYNAME):
    returnMonitorValue = ''
    mFilterPerformanceMointorList = [f for f in m if (f[u'DISPLAYNAME'] == DISPLAYNAME)]
    if len(mFilterPerformanceMointorList) > 0:
        mFilterPerformanceMointor = mFilterPerformanceMointorList[0]
        if u'data' in mFilterPerformanceMointor:
            returnMonitorValue = int(mFilterPerformanceMointor[u'data'][0].get(u'value'))
        else:
            returnMonitorValue = 0
    else:
        returnMonitorValue = 0
    return returnMonitorValue


def _get_one_performance_monitor_with_name(DISPLAYNAME):
    returnMonitorValue = ''
    mFilterPerformanceMointorList = [f for f in m if (f[u'DISPLAYNAME'] == DISPLAYNAME)]
    if len(mFilterPerformanceMointorList) > 0:
        mFilterPerformanceMointor = mFilterPerformanceMointorList[0]
        if u'data' in mFilterPerformanceMointor:
            returnMonitorValue = mFilterPerformanceMointor[u'DISPLAYNAME'] + ' --> ' + str(
                mFilterPerformanceMointor[u'data'][0].get(u'value'))
        else:
            returnMonitorValue = mFilterPerformanceMointor[u'DISPLAYNAME'] + '--> ......'
    else:
        returnMonitorValue = DISPLAYNAME + '--> ......'
    return returnMonitorValue


def _get_row_field(mList, filedName, filedValueEqual):
    filedFilterList = [f for f in mList if (filedValueEqual in f[filedName])]
    return filedFilterList

def load_config_yaml():
    f = open('config.yaml')
    setting = yaml.load(f)
    apiKeyConfig = setting['apiKey']
    opmUrlConfig = setting['opmUrl']
    return apiKeyConfig,opmUrlConfig


if __name__ == "__main__":
    #为输入参数，可以做到配置文件中
    apiKey,opmUrl = load_config_yaml()

    #apiKey = '5a1f39b05a4f2708b048b3852f12f7ca'
    #opmUrl = 'http://60.190.251.203:12390'
    category =['Server']

    #apiKey = 'f5334133f4f90bcd7a5a3d2c95cb22b6'
    #opmUrl = 'http://demo.opmanager.com:80'
    #category =['Server']

    outReportFileName ='opm_report_of_'+  datetime.datetime.now().strftime('%Y%m%d_%H%M_%S') #生成文件名称
    tableTitle = u'服务器主要性能指标一览表'   #表格名称
    title = [u'No.', u"主机名", u"类型", u"IP地址", u"CPU", u"内存", u"磁盘利用率"] #表格字段名称
    data = []
    row = []
    rowNumber = 0
    deviceName = ""  # 保留用的设备名称

    #根据分类，获取设备列表
    s = requests.Session()
    url = opmUrl+ '/api/json/device/listDevices?apiKey='+apiKey+'&category='+category[0]
    sList = s.get(url).json()

    if sList and isinstance(sList, list):#判断是否有结果集合
        for _s in sList:
            rowNumber = rowNumber + 1 #生成记录号
            deviceName = _get_value(_s, 'deviceName')
            row = [rowNumber, _get_value(_s, 'displayName'), _get_value(_s, 'type'), _get_value(_s, 'ipaddress')]
            #获取某一设备对应的指标列表，结果是一个LIST
            url2 = opmUrl+ '/api/json/device/getAssociatedMonitors?apiKey='+apiKey+'&name=' + _get_value(
                _s, 'deviceName')
            m = s.get(url2).json().get('performanceMonitors').get('monitors')
            if isinstance(m, list) and len(m) > 1:
                # 获取CPU利用率
                cpuUsage = _get_one_performance_monitor(u'CPU利用率')
                #如果没有返回结果或者返回结果为0，表格中置空
                if cpuUsage>0:
                    row.append(cpuUsage)
                else:
                    row.append('')

                #如果没有返回结果或者返回结果为0，表格中置空
                # 获取内存利用率
                memUsage = _get_one_performance_monitor(u'内存利用率')
                if memUsage>0:
                    row.append(memUsage)
                else:
                    row.append('')


                # 获取总的磁盘利用率
                diskUsage = _get_one_performance_monitor(u'磁盘利用率')

                # 获取磁盘分区信息，先过滤得到有设备分区信息的字段，WIN和LINUX不一样，然后进行排序
                perDiskUsage = ''
                mFilter = sorted([formatCell for formatCell in m if (
                    (formatCell[u'DISPLAYNAME'].find(u'设备分区信息(%)') != -1) or (formatCell[u'DISPLAYNAME'].find(u'设备的分区明细(%)') != -1))])
                for r in mFilter:
                    if u'data' in r:

                        # 处理分区标签，WIN下需要进一步处理掉“ Label:  Serial Number”，特别注意split函数是返回一个LIST的
                        diskLable = r[u'DISPLAYNAME'].split("(%)-")[1]
                        if diskLable.find('\ Label:  Serial Number'):
                            diskLable = diskLable.split("\\")[0]
                        dataMonitor = r.get(u'data')
                        for d in dataMonitor:
                            perDiskUsage = perDiskUsage + diskLable + " --> " + str(d.get(u'value')) + '%\n'
                #如果没有返回结果或者返回结果为0，表格中置空
                if diskUsage>0:
                    diskUsage = u'总磁盘利用率-->' + str(diskUsage) + u'%\n各分区利用率如下：\n' + perDiskUsage
                    row.append(diskUsage)
                else:
                    row.append("")
            else:
                row.append("")
                row.append("")
                row.append("")
            # print row
            data.append(row)
    #print data

    #创建EXCEL文件名称，并创建一个表
    workbook = xlsxwriter.Workbook(outReportFileName+'.xlsx')
    worksheet1 = workbook.add_worksheet(U'主要性能指标一览表')

    # 表头格式.
    formatTitle = workbook.add_format({
        'bg_color': '#cccccc',
        'align': 'left',
        'align': 'top',
        'bold': True,
        'text_wrap': True,
        'font_color': '#000000'})
    formatTitle.set_font_size(9)
    formatTitle.set_border(1)
    formatTitle.set_top()
    formatTitle.set_left()

    #表格单元格格式
    formatCell = workbook.add_format({'bg_color': '#FFFFFF',
                             'align': 'left',
                             'align': 'top',
                             'bold': False,
                             'text_wrap': True,
                             'font_color': '#000000'})
    formatCell.set_font_size(9)
    formatCell.set_border(1)
    formatCell.set_top()
    formatCell.set_align('left')


    #worksheet1.set_row(0, 40, cell_format)
    worksheet1.set_column('A:A',3 )
    worksheet1.set_column('B:B',18)
    worksheet1.set_column('C:D',12)
    worksheet1.set_column('E:F',4)
    worksheet1.set_column('G:G',18)
    worksheet1.set_row(0, 20)

    #表格大标题格式设置
    tableTitle_format = workbook.add_format({
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter',
        'fg_color': '#F8F8FF'})
    worksheet1.merge_range('A1:G1', tableTitle, tableTitle_format)
    #worksheet1.write_rich_string()
    worksheet1.write_row('A2', title, formatTitle)

    #单元格内容填充
    for rowid, row_data in enumerate(data):
        worksheet1.write_row(rowid + 2, 0, row_data, formatCell)

    #表格左注脚格式
    reportDateCellLeft ='A'+str(rowNumber+3)+':C'+str(rowNumber+3)
    reportDateCellRight ='D'+str(rowNumber+3)+':G'+str(rowNumber+3)

    footer_format_left = workbook.add_format({
        'bold': 0,
        'size':9,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter',
        'font_color': '#8A8A8A',
        'bg_color': '#cccccc'}

        )
    footer_format_right = workbook.add_format({
        'bold': 0,
        'size':9,
        'border': 1,
        'align': 'right',
        'valign': 'vcenter',
        'font_color': '#8A8A8A',
        'bg_color': '#cccccc'
    }
        )

    worksheet1.merge_range(reportDateCellLeft, u'备注：指标超过70%红色告警', footer_format_left)
    worksheet1.merge_range(reportDateCellRight, u'制表时间：'+datetime.datetime.now().strftime('%Y%m%d %H:%M %S'), footer_format_right)

    # Add a format. Light red fill with dark red text.
    formatAlarmCell = workbook.add_format({
                                   'font_color': '#FF4500'})
    # Write a conditional format over a range.
    alarmCell= 'E3:G'+str(rowNumber+3)
    worksheet1.conditional_format('E3:F17', {'type': 'cell',
                                             'criteria': '>=',
                                             'value': 75,
                                             'format': formatAlarmCell})

    workbook.close()
