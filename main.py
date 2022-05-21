import TEK_CSV

if __name__ == '__main__':
    mac_m1 = True

    get_test_info = True
    csv_to_excel = True
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
    add_file_name = True
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

    tek = TEK_CSV.tekCsv(csv_path=path_csv, excel_path=path_excel, filter_factor=LPF_factor, LPF=lpf, FFT=fft,
                         time_graph=graph_time, fft_graph=graph_FFT, time_window=10000, time_window_type='crop',
                         fft_window=5000, fft_window_type='crop')

    if add_info_file:
        tek.add_info_file(path_information)

    if csv_to_excel:
        tek.csv_to_excel(graph_time, graph_FFT)
        print('end csv_to_excel')