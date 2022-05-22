import os

import FILE_MANAGEMENT
import TEK_CSV
import GET_SUMMARY
# import name_change

if __name__ == '__main__':
    mac_m1 = True

    get_test_info = False
    csv_to_excel = False
    # csv_to_excel = False
    # add_info_file = True
    add_info_file = False
    if csv_to_excel:
        fft = True
        graph_time = True
        graph_FFT = True
        get_period = False
        lpf = False
        LPF_factor = 0.5
        get_period = False
        graph_FFT = True
    change_file_name = False
    get_summary = True
    if get_summary:
        get_by_option = True

    crop_time_window = 10000
    crop_fft_window = 5000

    if mac_m1:
        path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/sample/'
        path_csv = path + 'csv/'
        path_excel = path + 'excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test infomation/'

    fm = FILE_MANAGEMENT.FILE_MANAGEMENT()
    if csv_to_excel:
        tek = TEK_CSV.tekCsv(csv_path=path_csv, excel_path=path_excel, filter_factor=LPF_factor, LPF=lpf, FFT=fft,
                             time_graph=graph_time, fft_graph=graph_FFT, time_window=10000, time_window_type='crop',
                             fft_window=5000, fft_window_type='crop')
        tek.csv_to_excel(graph_time, graph_FFT)
        print('end csv_to_excel')

    if add_info_file:
        tek.add_info_file(path_information)

    # # name변경 함수 수정할 것

    if get_summary:
        excel_list = os.listdir(path_excel)
        excel_list = [file for file in excel_list if file[:3] == 'tek' and file.endswith('.xlsx')]

        summary = GET_SUMMARY.Get_Summary(excel_list, path_excel)
        summary.get_summary()