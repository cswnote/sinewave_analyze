import os
import pandas as pd
import FILE_MANAGEMENT
import TEK_CSV
import GET_SUMMARY
# import name_change

if __name__ == '__main__':
    mac_m1 = False

    get_test_info = False
    csv_to_excel = False
    # csv_to_excel = False
    # add_info_file = True
    add_info_file = False
    kmon_csv = True

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
    get_summary = False
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
    else:
        path = 'D:/winston/OneDrive - (주)필드큐어/data_analyze/'
        path_csv = path + 'csv/auto mode test/'
        path_excel = path + 'excel/auto mode test/'
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

    if kmon_csv:
        path_kmon = path_excel[:path_excel.find('excel')] + 'kmon_csv/'
        csv_list = os.listdir(path_kmon)
        csv_list = [file for file in csv_list if file[:10] == 'info_test_' and file.endswith('.csv')]
        # kmon_columns_name = [['Control', 'RF Volt Set Ch 1', 'RF Volt Set Ch 2', 'RF Volt Set Ch 3', 'RF Volt Set Ch 4',
        #                      'RF Curr Set Ch 1', 'RF Curr Set Ch 2', 'RF Curr Set Ch 3', 'RF Curr Set Ch 4',
        #                      'CP Pwm Set Ch 1', 'CP Pwm Set Ch 2', 'CP Pwm Set Ch  3', 'CP Pwm Set Ch 4', 'RF Pwm Set'],
        #                      ['RF Volt Ch 1', 'RF Volt Ch 2', 'RF Volt Ch 3', 'RF Volt Ch 4',
        #                      'RF Curr Ch 1', 'RF Curr Ch 2', 'RF Curr Ch 3', 'RF Curr Ch 4',
        #                      'CP Pwm Ch 1', 'CP Pwm Ch 2', 'CP Pwm Ch 3', 'CP Pwm Ch 4',
        #                      'Loop Time 0.1 us']]
        merge_kmon = GET_SUMMARY.Get_Summary(csv_list, path_kmon)
        kmon_set_file = path_kmon + 'kmon_monitoring_set.xlsx'

        df = merge_kmon.combine_kmon_data(kmon_set_file)

        filename = 'kmon_all.xlsx'
        if not os.path.exists(path_kmon + filename):
            with pd.ExcelWriter(path_kmon + filename, mode='w', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='total')
        else:
            with pd.ExcelWriter(path_kmon + filename, mode='a', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='total')