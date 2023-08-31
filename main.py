import os
import pandas as pd
import FILE_MANAGEMENT
import TEK_CSV
import GET_SUMMARY
import platform
import socket
# import name_change

if __name__ == '__main__':
    if 'Rainys-MacBook-Air-2' in socket.gethostname():
        path = os.getcwd() + '/data/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test_info/'
        path_kmon = path + 'kmon_csv/'

    elif 'SX70' in socket.gethostname():
        path = 'C:/data/PL150/RFAMP Voltage Current Accuracy/1st_correction_function/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test_info/'
        path_kmon = path + 'kmon_csv/'

    evaluation_control_file = 'eval_control.xlsx'
    kmon_sheet = 'kmon monitoring set RFamp'
    # kmon_sheet = 'kmon monitoring set main'

    # get_test_info = False
    csv_to_excel = False
    add_info_file = False
    change_file_name = False # # RFAMP 보드 전용


    if csv_to_excel:
        fft = False
        graph_time = False
        graph_FFT = False
        lpf = False
        LPF_factor = 0.5
        get_period = False

    get_summary = True
    kmon_csv = True

    seperate_data_by_tag = False
    if get_summary:
        get_by_option = False

    # # time_window_type is 'ratio' than the graph x numbers is (data_length / time_window)
    crop_time_window = 1
    crop_fft_window = 500 # 126

    fm = FILE_MANAGEMENT.FILE_MANAGEMENT()
    if csv_to_excel:
        tek = TEK_CSV.tekCsv(path=path, time_window_type='ratio', time_window=crop_time_window, LPF=lpf, filter_factor=LPF_factor,
                             fft_window_type='crop', fft_window=crop_fft_window, graph_time=graph_time, graph_FFT=graph_FFT)
        tek.csv_to_excel()
        print('end csv_to_excel')

    if change_file_name:
        tek = TEK_CSV.tekCsv(path=path, eval_fila=evaluation_control_file)
        sheet = 'file name'
        tek.file_name_change(sheet)

    if add_info_file:
        tek = TEK_CSV.tekCsv(path=path, time_window_type='ratio', time_window=1,
                             fft_window=1, fft_window_type='crop', fft_window_size=1000)
        tek.combine_infofiles()

    # # name변경 함수 수정할 것

    if get_summary:
        excel_list = os.listdir(path_excel)
        excel_list = [file for file in excel_list if file[:3] == 'tek' and file.endswith('.xlsx')]

        summary = GET_SUMMARY.Get_Summary(path, evaluation_control_file)
        summary.get_summary()

    if kmon_csv:
        path_kmon = path_excel[:path_excel.find('tek_excel')] + 'kmon_csv/'
        csv_list = os.listdir(path_kmon)
        csv_list = [file for file in csv_list if file[:10] == 'info_kmon_' and file.endswith('.csv')]
        csv_list.sort()

        merge_kmon = GET_SUMMARY.Get_Summary(path, evaluation_control_file)

        test_files = pd.read_excel(path + evaluation_control_file, sheet_name='file name')
        test_files = test_files.iloc[:, 0].tolist()

        for file in test_files:
            df = merge_kmon.combine_kmon_data(file, kmon_sheet)
            merge_kmon.check_kmon_and_testfile(df, file)
        merge_kmon.merge_kmon_and_summary()

    if seperate_data_by_tag:
        merge_kmon = GET_SUMMARY.Get_Summary(path, evaluation_control_file)
        summary = 'summary.xlsx'

        merge_kmon.get_seperated_data(summary)