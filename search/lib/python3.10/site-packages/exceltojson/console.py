import sys
import getopt


from exceltojson.excel2json import ProcessExcel

# __all__ = ['ProcessExcel', 'main', 'usage']


def usage():
    print("""
-h | --help: get help document
-S | --notShowRow: line number to key
-s | --sourcePath: excel file path
-o | --outDir: json file save dir
-P | --noPatchAlias: usr header alias, if no alias use column header as the key.
-M | --noMergeCell: if empty cell, use as merge cell, the content will be same with above cell.
-r | --rowMax:  default 1000, type int, to use this limit json file size
-i | --index: sheet index list , eg: -i 0, 1, 2
-n | --names: sheet name list, eg: -n name1,name2,name3
-a | --alias: change column header name and as this key word, eg: -a header1:alias1,header2:alias2;otherHeader:otherAlias

note: -a and (-i or -n) must in pairs
""")


def main():
    try:
        opts, args = getopt.getopt(
            sys.argv[1:],
            "hMPr:a:i:n:o:s:S",
            ["help", "rowMax", "noMergeCell", "noPatchAlias",
             "alias", "index", "names", "outDir", "sourcePath", "noShowRow"])
    except getopt.GetoptError as e:
        print(str(e))
        print('use -h or --help to get help')
        sys.exit(-1)

    excel_path = ''
    output_dir = ''
    row_max = 1000
    merge_cell = True
    patch_alias = True
    show_row = True
    index = []
    names = []
    alias = []

    for o, a in opts:
        if o in ('-S', '--noShowRow'):
            show_row = False
        if o in ('-s', '--sourcePath'):
            excel_path = a
        elif o in ("-h", "--help"):
            usage()
            sys.exit()
        elif o in ("-o", "--outDir"):
            output_dir = a
        elif o in ('-P', '--noPatchAlias'):
            patch_alias = False
        elif o in ('-M', '--noMergeCell'):
            merge_cell = False
        elif o in ('-r', '--rowMax'):
            try:
                row_max = int(a)
            except ValueError:
                print('-r, --rowMax should be a integer value')
                sys.exit(-1)
        elif o in ('-i', '--index'):
            temp = a.split(',')
            try:
                [int(i) for i in temp]
            except ValueError:
                print('-i, --index should be a string that comma separated values, '
                      'the value each one should be a int value, like (-i 0,1,2) ')
                sys.exit(-1)
            index = temp
        elif o in ('-n', '--names'):
            names = a.split(',')
        elif o in ('-a', '--alias'):
            temp = a.split(';')
            for t in temp:
                temp_dict = {}
                for d in t.split(','):
                    data = d.split(':', 1)
                    if len(data) != 2:
                        print('-a, --alias should be a string that semicolon separated values, '
                              'the value is a comma separated each one should contain a colon separated char, like'
                              '(-a header1:alias1,header2:alias2;otherHeader:otherAlias)')
                        sys.exit(-1)
                    temp_dict[data[0].strip()] = data[1].strip()
                alias.append(temp_dict)

    if output_dir and excel_path is False:
        print('output directory and excel source file should be have')
        sys.exit(-1)

    # if alias paris with index or names
    alias_desc = '(-a, --alias) value must in pairs with (-i, --index) value or (-n, --names) value'

    _exit = False

    if alias:
        if index:
            if len(index) != len(alias):
                _exit = True
        elif names:
            if len(names) != len(alias):
                _exit = True
        else:
            _exit = True

    elif not alias:
        if index:
            _exit = True
        elif names:
            _exit = True

    if _exit:
        print(alias_desc)
        sys.exit(-1)

    def get_pairs(_list):
        return {key: value for key, value in zip(_list, alias)}

    try:
        if index:
            pairs = get_pairs(index)
            ProcessExcel(excel_path, output_dir, pairs, None, merge_cell, show_row, patch_alias)(row_max)
        elif names:
            pairs = get_pairs(names)
            ProcessExcel(excel_path, output_dir, None, pairs, merge_cell, show_row, patch_alias)(row_max)
        else:
            ProcessExcel(excel_path, output_dir, None, None, merge_cell, show_row, patch_alias)(row_max)
    except ValueError as e:
        print(str(e))

# if __name__ == '__main__':
#     main()
