import os
import pandas as pd
import FILE_MANAGEMENT
import TEK_CSV
import GET_SUMMARY
import platform
# import name_change

if __name__ == '__main__':
    if platform.platform()[:3].lower() == 'mac':
        mac_m1 = True
    elif platform.platform()[:3].lower() == 'win':
        mac_m1 = False

    evaluation_control_file = 'eval_control.xlsx'

    # get_test_info = False
    csv_to_excel = False
    # csv_to_excel = False
    add_info_file = False
    # add_info_file = False
    change_file_name = False


    if csv_to_excel:
        fft = True
        graph_time = True
        graph_FFT = True
        lpf = False
        LPF_factor = 0.5
        get_period = False
        graph_FFT = True

    get_summary = False
    kmon_csv = False
    if get_summary:
        get_by_option = False

    crop_time_window = 10000
    crop_fft_window = 5000

    if mac_m1:
        path = '/Users/rainyseason/winston/Workspace/python/Pycharm Project/sinewave_analyze/Evaluation/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test information/'
        path_kmon = path + 'kmon_csv'
    else:
        path = 'D:/data_analyze/'
        path_csv = path + 'tek_csv/'
        path_excel = path + 'tek_excel/'
        path_summary = path + 'summary/'
        path_information = path + 'test information/'
        path_kmon = path + 'kmon_csv/'

    fm = FILE_MANAGEMENT.FILE_MANAGEMENT()
    if csv_to_excel:
        tek = TEK_CSV.tekCsv(path=path, time_window_type='ratio', time_window=10,
                             fft_window_type='crop', fft_window=1000)
        tek.csv_to_excel(graph_time, graph_FFT)
        print('end csv_to_excel')

    if change_file_name:
        tek = TEK_CSV.tekCsv(path=path, eval_fila=evaluation_control_file)
        sheet = 'file name'
        tek.file_name_change(sheet)

    if add_info_file:
        tek = TEK_CSV.tekCsv(path=path, time_window_type='crop', time_window=10000,
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
            df = merge_kmon.combine_kmon_data(file)
            merge_kmon.check_kmon_and_testfile(df, file)
        merge_kmon.merge_kmon_and_summary()

    merge_kmon = GET_SUMMARY.Get_Summary(path, evaluation_control_file)
    summary = 'summary 66 ~ 73 AMP03 300ohm.xlsx'
    merge_kmon.get_seperated_data(summary)

