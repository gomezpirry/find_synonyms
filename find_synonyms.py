import obonet
import networkx
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

    # for each obo file
    for i in files_read:
        print('read file ', i)
        # read .obo file
        graph = obonet.read_obo(i)
        print(len(graph), 'nodes')

        # get all nodes of graph
        nodes = graph.nodes(data=True)

        # get ids
        ids = {}
        for cell_ids in ids_column:
            print('check id:  ', str(int(cell_ids.value)))
            # for all nodes in graph
            for node in nodes:
                # find id node corresponding to excel id
                if str(node[0]) == str(int(cell_ids.value)):
                    # get synonyms
                    ids[node[0]] = node[1]['synonym']
                # end if
            # end for
        # end for
        style = xlwt.easyxf('pattern: pattern solid, fore_colour red;')
        for key, values in ids.items():

            for id_cell in range(len(ids_column)):
                # find id cell corresponding to results
                if str(int(ids_column[id_cell].value)) == str(key):
                    write_sheet.write(id_cell + index_row + 1, index_col + 1, len(values))
                    # write all synonyms
                    current_index = 2
                    for synonym in range(len(values)):
                        write_sheet.write(id_cell + index_row + 1, index_col + current_index, "synonym : " + str(synonym + 1), style = style)
                        current_index += 1
                        synonym_split = values[synonym].split('"', 2)
                        # get synonym name
                        synonym_name = synonym_split[1]
                        write_sheet.write(id_cell + index_row + 1, index_col + current_index, synonym_name)
                        current_index += 1
                        # get synonym type
                        synonym_rest_split = synonym_split[2].split('[')
                        synonym_type = synonym_rest_split[0]
                        write_sheet.write(id_cell + index_row + 1, index_col + current_index, synonym_type)
                        current_index += 1
                        # get synonym config
                        synonym_config = str(synonym_rest_split[1])[0:str(synonym_rest_split[1]).find("{")]
                        synonym_config_value = str(synonym_rest_split[1])[str(synonym_rest_split[1]).find("{"):
                                                                          str(synonym_rest_split[1]).find("}")].split(',')
                        write_sheet.write(id_cell + index_row + 1, index_col + current_index, synonym_config)
                        current_index += 1
                        # for all config values of synonym
                        for synonym_part in range(len(synonym_config_value)):
                            write_sheet.write(id_cell + index_row + 1, index_col + current_index,
                                              synonym_config_value[synonym_part])
                            current_index += 1


    if output == '':
        write_book.save(file_input_xls)
    else:
        write_book.save(output)

# end main


if __name__ == "__main__":
    main(sys.argv[1:])
