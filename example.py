from excel_table import ExcelTable

def main():

    serialized_data = [
        {
            'alpha': {
                'quebec': True,
                'papa': {
                    'romeo': 'Alabama',
                    'sierra': 'Georgia',
                }
            },
            'bravo': 2,
            'charlie': 3,
        },
        {
            'alpha': null,
            'bravo': 5,
            'charlie': 6,
        },
    ]

    workbook = xlsxwriter.Workbook('example.xlsx')

    excel_table = ExcelTable(
        data=serialized_data,
        columns=dict(
            quebec='alpha.quebec',
            romeo='alpha.papa.romeo',
            sierra='alpha.papa.sierra',
            bravo=None,
            charlie=None,
            delta=dict(
                data_accessor=lambda item: None,
                formula='=AVERAGE({bravo}, {charlie})',
            ),
        )
    )

    worksheet = workbook.add_worksheet()

    worksheet.add_table(
        *excel_table.coordinates,
        {
            'name': 'Table1',
            'style': 'Table Style Light 15',
            'columns': excel_table.columns,
            'data': excel_table.data,
            'total_row': True,
        }
    )

    workbook.close()

    print('Done')

if __name__ == '__main__':
    main()
