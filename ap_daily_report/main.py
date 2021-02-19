import read_csv as rc
import re


def process():
    print('file name :', end='')
    file_name = input()
    wb = rc.load_wb(file_name)

    # get sheet name(+result)
    sheet_names_list = rc.get_sheet_names_list(wb)

    # get right sheet name(compare A) ret type : list
    list_a_sheet_name = ''.join(sheet_names_list[-1])

    # delete 'result' sheet
    p = re.compile('result*')
    for sheet_name in sheet_names_list:
        if p.match(sheet_name):
            rc.delete_result_sheet(wb, sheet_name)
            list_a_sheet_name = ''.join(sheet_names_list[-2])

    # get hostnames in right sheet
    list_a_hostnames = rc.read_xlsx(wb, list_a_sheet_name)
    if list_a_hostnames is None:
        print('faillllllllllllllllll')
        wb.close()
        exit()

    # create 'result' sheet
    rc.create_result(wb, list_a_hostnames)

    # get sheet name(-result)
    sheet_names_list = rc.get_sheet_names_list(wb)
    for i in range(0, len(sheet_names_list) - 2):
        # get hostnames in left sheet
        list_b_hostnames = rc.read_xlsx(wb, sheet_names_list[i])
        compare(wb, list_a_hostnames, list_b_hostnames)

    rc.save_wb(wb, file_name)
    wb.close()
    print('complete!')


def compare(wb, list_a, list_b):
    for i in list_a:
        if i in list_b:
            rc.count_cnt(wb, i)


if __name__ == '__main__':
    process()

