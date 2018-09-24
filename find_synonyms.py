import obonet
import networkx
import xlrd
import xlwt
import sys
import os
import getopt
from xlutils.copy import copy


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

    if not output == '':
        file_input_xls = os.path.basename(input_xls)
        filename_xls, file_extension_xls = os.path.splitext(file_input_xls)

    # verify if input obo file exist
    if not os.path.isfile(file_input_obo):
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

    # read .obo file
    graph = obonet.read_obo(url)
    print(len(graph))

    # get all nodes of graph
    nodes = graph.nodes(data=True)

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
            print(cell, row[cell])
            if row[cell].value == 'ID':
                index_col = cell
                find_id_col = True
                break
            # end if
        # end for
        if index_row > 10:
            print("ID cell not found")
            break
        if not find_id_col:
            index_row += 1
    # end while
    print(index_row, index_col)

    # get ids
    ids = {}
    # get column later ID
    ids_column = sheet.col(index_col)[index_row + 1:]
    for cell_ids in ids_column:
        # for all nodes in graph
        for node in nodes:
            # find id node corresponding to excel id
            if str(node[0]) == str(int(cell_ids.value)):
                # get synonyms
                ids[node[0]] = node[1]['synonym']
            # end if
        # end for
    # end for

    for key, values in ids.items():

        for id_cell in range(len(ids_column)):
            # find id cell corresponding to results
            if str(int(ids_column[id_cell].value)) == str(key):
                write_sheet.write(id_cell + index_row + 1, index_col + 1, len(values))
                # write all synonyms
                for synonym in range(len(values)):
                    write_sheet.write(id_cell + index_row + 1, index_col + synonym + 2, values[synonym])

    if output == '':
        write_book.save(file_input_xls)
    else:
        write_book.save(output)
    print(ids)
# end main


if __name__ == "__main__":
    main(sys.argv[1:])
