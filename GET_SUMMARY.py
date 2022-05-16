import openpyxl
import os
import numpy as np
from win32com.client import Dispatch




class Get_Summary():
    def __init__(self):
        super().__init__()

    def get_summary(self, file_list, path):
        excel_list = file_list

        summary_wb = openpyxl.Workbook()
        summary_ws = summary_wb.active
        summary_ws.title = 'summary'

        # summary_ws('a1').value = excel_file.spilt('.xlsx')[0]

        # summary_ws['a1'].value = 'filename'
        # summary_ws['b1'].value = 'V Frequency[MHz]'
        # summary_ws['c1'].value = 'Delay(degree)'
        # summary_ws['d1'].value = 'Ave. RP Coff'
        # summary_ws['e1'].value = 'Vrms'
        # summary_ws['f1'].value = 'Irms'
        # summary_ws['g1'].value = 'Real P[W]'
        # summary_ws['h1'].value = 'V dc'
        # summary_ws['i1'].value = 'I dc'
        # summary_ws['j1'].value = 'FFT V freq[MHz]'
        # summary_ws['k1'].value = 'FFT V rms'
        # summary_ws['l1'].value = 'FFT V dc abs'
        # summary_ws['m1'].value = 'FFT I freq[MHz]'
        # summary_ws['n1'].value = 'FFT I rms'
        # summary_ws['o1'].value = 'FFT I dc abs'

        for idx, excel_file in enumerate(excel_list):
            print('in summary process: ', idx + 1, '/', len(excel_list), '    ', excel_file)
            wb = openpyxl.load_workbook(path + excel_file)
            try:
                ws = wb[excel_file.split(' ')[0]]
            except:
                ws = wb[excel_file.split('.')[0]]
            summary_ws.cell(idx + 2, 1).value = excel_file.split('.xlsx')[0]

            file_space = len(excel_file.split('.xlsx')[0].split(' '))

            summary_name = False
            if not summary_name:
                summary_ws['a1'].value = 'filename'
                summary_ws.cell(1, file_space + 1).value = 'V Frequency[MHz]'
                summary_ws.cell(1, file_space + 2).value = 'Delay(degree)'
                summary_ws.cell(1, file_space + 3).value = 'Ave. RP Coff'
                summary_ws.cell(1, file_space + 4).value = 'Vrms'
                summary_ws.cell(1, file_space + 5).value = 'Irms'
                summary_ws.cell(1, file_space + 6).value = 'Real P[W]'
                summary_ws.cell(1, file_space + 7).value = 'Vmean'
                summary_ws.cell(1, file_space + 8).value = 'Imean'
                summary_ws.cell(1, file_space + 9).value = 'FFT V freq[MHz]'
                summary_ws.cell(1, file_space + 10).value = 'FFT V rms'
                summary_ws.cell(1, file_space + 11).value = 'FFT V dc abs'
                summary_ws.cell(1, file_space + 12).value = 'FFT I freq[MHz]'
                summary_ws.cell(1, file_space + 13).value = 'FFT I rms'
                summary_ws.cell(1, file_space + 14).value = 'FFT I dc abs'
                summary_name = True


            for i in range(file_space):
                summary_ws.cell(idx + 2, i + 1).value = excel_file.split('.xlsx')[0].split(' ')[i]
                if i == file_space - 2:
                    summary_ws.cell(idx + 2, i + 1).value = float(
                        (excel_file.split('.xlsx')[0].split(' ')[i]).split('ohm')[0])
                if i == file_space - 1:
                    summary_ws.cell(idx + 2, i + 1).value = int((excel_file.split('.xlsx')[0].split(' ')[i]).split('PWM')[1])

            # # Delay(degree)
            for i in range(100, 8, -1):
                if ws.cell(5, i).value is not None:
                    summary_ws.cell(idx + 2, file_space + 2).value = ws.cell(5, i).value
                    break

            # # V freq
            for i in range(100, 8, -1):
                if ws.cell(1, i).value is not None:
                    summary_ws.cell(idx + 2, file_space + 1).value = ws.cell(1, i).value
                    break

            # # Ave. RP Coff
            summary_ws.cell(idx + 2, file_space + 3).value = ws.cell(7, 8).value

            # # Vrms
            for i in range(100, 8, -1):
                if ws.cell(8, i).value is not None:
                    summary_ws.cell(idx + 2, file_space + 4).value = ws.cell(8, i).value
                    break

            # # Irms
            for i in range(100, 8, -1):
                if ws.cell(9, i).value is not None:
                    summary_ws.cell(idx + 2, file_space + 5).value = ws.cell(9, i).value
                    break

            # # real power
            for i in range(100, 8 , -1):
                if ws.cell(10, i).value is not None:
                    summary_ws.cell(idx + 2, file_space + 6).value = ws.cell(10, i).value
                    break

            # # Vdc
            summary_ws.cell(idx + 2, file_space + 7).value = ws['f1'].value

            # # Idc
            summary_ws.cell(idx + 2, file_space + 8).value = ws['f3'].value

            # # get FFT info

            v_flag = False
            i_flag = False
            for i in range(9, 1, -1):
                if ws.cell(13, i).value == 'V' and not v_flag:
                    ws_fft = wb['FFT_' + ws.cell(21, i).value]
                    try:
                        summary_ws.cell(idx + 2, file_space + 9).value = float(ws_fft['f1'].value / 10**6)
                    except:
                        summary_ws.cell(idx + 2, file_space + 9).value = ws_fft['f1'].value
                    summary_ws.cell(idx + 2, file_space + 10).value = ws_fft['f2'].value
                    try:
                        summary_ws.cell(idx + 2, file_space + 11).value = ws_fft['c2'].value / np.sqrt(2)
                    except:
                        summary_ws.cell(idx + 2, file_space + 11).value = ws_fft['c2'].value
                    v_flag = True
                elif ws.cell(13, i).value == 'A' and not i_flag:
                    ws_fft = wb['FFT_' + ws.cell(21, i).value]
                    try:
                        summary_ws.cell(idx + 2, file_space + 12).value = float(ws_fft['f1'].value / 10**6)
                    except:
                        summary_ws.cell(idx + 2, file_space + 12).value = ws_fft['f1'].value
                    summary_ws.cell(idx + 2, file_space + 13).value = ws_fft['f2'].value
                    try:
                        summary_ws.cell(idx + 2, file_space + 14).value = ws_fft['c2'].value / np.sqrt(2)
                    except:
                        summary_ws.cell(idx + 2, file_space + 14).value = ws_fft['c2'].value
                elif v_flag and i_flag:
                    break

            wb.close()

        summary_wb.save(path + 'summary.xlsx')
        summary_wb.close()

    def copy_paste_graph(self, **kwargs):
        path = kwargs.get('path', os.getcwd() + '\\')
        file_orders = kwargs.get('file_list', 'Sheet1')
        summary = kwargs.get('summary_file_name', 'summary')

        file_list = os.listdir(path)
        file_list = [file for file in file_list if file.endswith(".xlsx")]

        for list in file_list:
            if list[0:3] != 'tek' or list[-4:] != 'xlsx':
                file_list.remove(list)

        summary_wb = openpyxl.load_workbook(path + summary + '.xlsx')
        order_ws = summary_wb[file_orders]

        graph_head = []

        for i in range(1, 16385):
            if order_ws.cell(1, i).value is not None:
                graph_head.append(order_ws.cell(1, i).value)
            else:
                break

        graph_folder = path + 'graph/'
        try:
            if not os.path.exists(graph_folder):
                os.makedirs(graph_folder)
                print("creat '%s' folder" % graph_folder)
        except OSError:
            print('Error: Creating directory. ' + graph_folder)


        for i in range(len(graph_head)):
            print(graph_head[i])
            graph_ws = summary_wb.create_sheet(graph_head[i])
            graph_list = []

            for j in range(2, 1048576):
                if order_ws.cell(j, i + 1).value is not None:
                    temp = order_ws.cell(j, i + 1).value.lower()
                    if len(temp[3:]) == 1:
                        temp = 'tek' + '000' + temp[3:]
                        graph_list.append(temp)
                    elif len(temp[3:]) == 2:
                        temp = 'tek' + '00' + temp[3:]
                        graph_list.append(temp)
                    elif len(temp[3:]) == 3:
                        temp = 'tek' + '0' + temp[3:]
                        graph_list.append(temp)
                    elif len(temp[3:]) >= 5:
                        print('graph file name error!!!', end='    ')
                        print(order_ws.cell(j, i + 1).value)
                    else:
                        graph_list.append(temp)
                else:
                    break

            absent_file = []
            for graph in graph_list:
                if graph + '.xlsx' not in file_list:
                    absent_file.append(graph)
                    graph_list.remove(graph)
            # graph_ws.cell(1, 1).value = absent_file
            for idx, file in enumerate(absent_file):
                graph_ws.cell(1, j + 2).value = file[idx]

            for idx, graph in enumerate(graph_list):
                # wb = openpyxl.load_workbook(path + graph + '.xlsx')
                # ws = wb[graph]
                #
                # num_of_data_len = int(ws['b10'].value)
                # if num_of_data_len > 1000000 - 21:
                #     num_of_data_len = 1000000 - 21
                #
                # for j in range(2, 100000):
                #     if ws.cell(13, j).value is None:
                #         max_cal = j - 1
                #         break
                #
                # chart1 = openpyxl.chart.LineChart()
                # chart1.title = graph
                # chart1.style = 10
                # chart1.x_axis.title = "time"
                #
                # chart2 = openpyxl.chart.LineChart()
                # chart2.y_axis.majorGridlines = None
                # chart2.y_axis.axId = 200
                #
                # cats = openpyxl.chart.Reference(ws, min_col=1, min_row=22, max_row=num_of_data_len + 21)
                # for j in range(max_cal, max_cal-2, -1):
                #     if ws.cell(13, j).value == 'V':
                #         data1 = openpyxl.chart.Reference(ws, min_col=j, max_col=j, min_row=21,
                #                                          max_row=num_of_data_len + 21)
                #         print(type(data1))
                #         chart1.add_data(data1, titles_from_data=True)
                #         chart1.set_categories(cats)
                #     elif ws.cell(13, j).value == 'A':
                #         data2 = openpyxl.chart.Reference(ws, min_col=j, max_col=j, min_row=21,
                #                                          max_row=num_of_data_len + 21)
                #         print(help(openpyxl.chart.Reference))
                #         chart2.add_data(data2, titles_from_data=True)
                #         chart2.set_categories(cats)
                #
                #
                #
                # s1 = chart1.series[0]
                # s1.graphicalProperties.line.width = 0
                #
                # s2 = chart1.series[0]
                # s2.graphicalProperties.line.width = 0
                #
                # chart2.y_axis.crosses = "max"  # max인 축이 오른쪽에 위치
                #
                # chart1 += chart2
                #
                # chart_location_idx = 35 * idx
                #
                # chart_location = 'b' + str(4 + chart_location_idx)
                # graph_ws.add_chart(chart1, chart_location)
                # chart1.chart_width = 30
                # chart1.chart_height = 18

                # wb.close()

                excel = Dispatch('Excel.Application')
                excel.Visible = True
                wb = excel.Workbooks.Open(path + graph + '.xlsx')
                sheet = wb.Worksheets(graph)
                mychart = sheet.ChartObjects(1)
                mychart.Chart.Export(Filename=graph_folder + str(idx + 1) + ' - ' + graph + '.jpg')
            excel.Quit()

            summary_wb.save(path + idx + ' - ' + summary + '.xlsx')



if __name__=='__main__':
    # path = os.getcwd() + '\\test\\'
    path = 'D:/work/data_analyze/'
    csv_path = path + 'csv/'
    print(path)

    summary_file = 'summary'
    summary = Get_Summary()

    summary.copy_paste_graph(path=path, summary_file_name=summary_file)
