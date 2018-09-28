import xlrd
import xlwt
import sys
import os
import getopt
from xlutils.copy import copy
import glob


def main(argv):

    # verify numbers of arguments
    if len(argv) < 4:
        print('Number of arguments invalid')
        print('Try with:')
        print('find_synonym.py -i <inputfile obo> -x <inputfile xls>')
        print('find_synonym.py -i <inputfile obo> -x <inputfile xls> -o <outputfile xls>')
        sys.exit()
    # end if

    url = ''
    input_xls = ''
    output = ''

    try:
        opts, args = getopt.getopt(argv, "hi:x:o:", ["ifile=", "xfile=", "ofile="])
    except getopt.GetoptError:
        print('read_synonym.py -i <inputfile obo> -x <inputfile xls>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print('read_synonym.py -i <inputfile> -x <inputfile xls>')
            sys.exit()
        elif opt in ("-i", "--ifile"):
            url = arg
        elif opt in ("-x", "--xfile"):
            input_xls = arg
        elif opt in ("-o", "--ofile"):
            output = arg
    # end for

    file_input_obo = os.path.basename(url)
    filename_obo, file_extension_obo = os.path.splitext(file_input_obo)

    file_input_xls = os.path.basename(input_xls)
    filename_xls, file_extension_xls = os.path.splitext(file_input_xls)

    files_read = []
    for filename in glob.glob(filename_obo + '*.obo'):
        files_read.append(filename)

    if not output == '':
        file_input_xls = os.path.basename(input_xls)
        filename_xls, file_extension_xls = os.path.splitext(file_input_xls)

    # verify if input obo file exist
    if len(files_read) == 0:
        print('Input .obo File does not exist')
        sys.exit()
    # end if

    # verify if input xls file exist
    if not os.path.isfile(file_input_xls):
        print('Input .xls File does not exist')
        sys.exit()
    # end if

    # verify if input file is a obo
    if not file_extension_obo == '.obo':
        print('Input File must have a .obo extension')
        sys.exit()
    # end if

    # verify if input file is a xls or xlsx
    if not (file_extension_xls == '.xls' or file_extension_xls == '.xlsx'):
        print('Input File must have a .xls or .xlsx extension')
        sys.exit()
    # end if

    # verify if output file is a xls
    if not output == '':
        if not (file_extension_xls == '.xls' or file_extension_xls == '.xlsx'):
            print('Input File must have a .xls or .xlsx extension')
            sys.exit()
        # end if
    # end if

    book = xlrd.open_workbook(input_xls)
    sheet = book.sheet_by_index(0)
    write_book = copy(book)
    write_sheet = write_book.get_sheet(0)

    # find ID col
    find_id_col = False
    index_col = 0
    index_row = 0
    while not find_id_col:
        row = sheet.row(index_row)
        for cell in range(len(row)):
            if row[cell].value == 'ID':
                index_col = cell
                find_id_col = True
                break
            # end if
        # end for
        if index_row > 10:
            "ID cell not found"
            break
        if not find_id_col:
            index_row += 1
    # end while
    print('ID in : ', index_row, index_col)

    # get column later ID
    ids_column = sheet.col(index_col)[index_row + 1:]
    style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')

    # for each obo file
    for i in files_read:
        print('read file ', i)
        count_lines = 0
        count_synonym = 0
        current_id_text = ""
        with open(i, encoding='utf-8') as infile:
            find_id = False
            count_write_synonym = 1
            for line in infile:
                line_split = line.split(' ')
                if 'id:' in line_split:
                    id_text = line.split(':')[1].replace('\n', '').strip()
                    for id_cell in range(len(ids_column)):
                        if str(ids_column[id_cell].value).split('.')[0].isnumeric():
                            if id_text == str(int(ids_column[id_cell].value)):
                                find_id = True
                                current_id_text = id_text
                                print(str('id :' + id_text))
                                continue
                if find_id:
                    if 'synonym:' in line_split:
                        count_synonym += 1
                        synonym_text = ' '.join(line_split)
                        for id_cell in range(len(ids_column)):
                            if str(ids_column[id_cell].value).split('.')[0].isnumeric():
                                if str(int(ids_column[id_cell].value)) == str(current_id_text):
                                    write_sheet.write(id_cell + index_row + 1, index_col + 1 + count_write_synonym,
                                                      "synonym : " + str(count_synonym), style=style)
                                    count_write_synonym += 1
                                    synonym_split = synonym_text.split('"', 2)
                                    if len(synonym_split) <= 2:
                                        print('bad write synonym')
                                    else:
                                        # get synonym name
                                        synonym_name = synonym_split[1]
                                        write_sheet.write(id_cell + index_row + 1, index_col + 1 + count_write_synonym,
                                                          synonym_name)
                                        count_write_synonym += 1
                                        synonym_rest_split = synonym_split[2].split('[')
                                        synonym_type = synonym_rest_split[0]
                                        write_sheet.write(id_cell + index_row + 1, index_col + 1 + count_write_synonym,
                                                          synonym_type)
                                        count_write_synonym += 1
                                        # get synonym config
                                        synonym_config = str(synonym_rest_split[1])[
                                                         0:str(synonym_rest_split[1]).find("{")]
                                        synonym_config_value = str(synonym_rest_split[1])[
                                                               str(synonym_rest_split[1]).find("{"):
                                                               str(synonym_rest_split[1]).find("}")].split(',')
                                        write_sheet.write(id_cell + index_row + 1, index_col + 1 + count_write_synonym,
                                                          synonym_config)
                                        count_write_synonym += 1
                                        # for all config values of synonym
                                        for synonym_part in range(len(synonym_config_value)):
                                            write_sheet.write(id_cell + index_row + 1, index_col + 1 + count_write_synonym,
                                                              synonym_config_value[synonym_part])
                                            count_write_synonym += 1
                                continue
                    term_text = line_split[0].replace('\n', '').strip()
                    if '[Term]' in term_text:
                        find_id = False
                        for id_cell in range(len(ids_column)):
                            if str(ids_column[id_cell].value).split('.')[0].isnumeric():
                                if str(int(ids_column[id_cell].value)) == str(current_id_text):
                                    write_sheet.write(id_cell + index_row + 1, index_col + 1, count_synonym)
                        count_synonym = 0
                        count_write_synonym = 1
                count_lines += 1
                print('file: ' + i + " -- " + str(count_lines))

    if output == '':
        write_book.save(file_input_xls)
    else:
        write_book.save(output)


if __name__ == "__main__":
    main(sys.argv[1:])





















