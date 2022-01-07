# xlsxwriter-tables
Easily export nested data to Excel

This class is intended to be used with [XlsxWriter](https://xlsxwriter.readthedocs.io/working_with_tables.html). It serves several purposes:

1. Co-location of column info and data-generation logic
2. Easily specify deeply nested data as the source for column data
3. Reference column headers dynamically in column formulas

The [example.py](example.py) file shows basic usage; [excel_table.py](excel_table.py) is also thoroughly documented. I intend to document more examples of usage in the future.

## API

<details>
    <summary>Sample data</summary>

    ``` py
    serialized_data = [
        {
            'alpha': {
                'oscar': True,
                'papa': {
                    'romeo': State(
                        name='Alabama', 
                        statehood_granted=date(1819, 12, 14),
                        symbols={
                            'bird': 'Yellowhammer',
                            'flower': 'Camellia',
                        },
                    ),
                    'sierra': State(
                        name='Georgia', 
                        statehood_granted=date(1788, 1, 2),
                        symbols={
                            'bird': 'Brown Thrasher',
                            'flower': 'Cherokee Rose',
                        },
                    ),
                }
            },
            'bravo': 22,
            'charlie': 4,
        },
        {
            'alpha': {
                'oscar': False,
                'papa': {
                    'romeo': State(
                        name='Minnesota', 
                        statehood_granted=date(1858, 5, 11),
                        symbols={
                            'bird': 'Common Loon',
                            'flower': 'Ladys Slipper',
                        },
                    ),
                    'sierra': State(
                        name='Wisconsin', 
                        statehood_granted=date(1848, 5, 29),
                        symbols={
                            'bird': 'Robin',
                            'flower': 'Wood Violet',
                        },
                    ),
                }
            },
            'bravo': 32,
            'charlie': 30,
        },
        {
            'alpha': {
                'oscar': None,
                'papa': {
                    'romeo': State(
                        name='Maryland', 
                        statehood_granted=date(1776, 7, 4),
                        symbols={
                            'bird': 'Baltimore Oriole',
                            'flower': 'Black-Eyed Susan',
                        },
                    ),
                    'sierra': State(
                        name='Virginia', 
                        statehood_granted=date(1788, 6, 25),
                        symbols={
                            'bird': 'Cardinal',
                            'flower': 'Flowering Dogwood',
                        },
                    ),
                }
            },
            'bravo': 7,
            'charlie': 10,
        },
    ]
    ```
    
</details>

Given the sample data above, we can generate an Excel table for export using XlsxWriter with the following code:

``` py
excel_table = ExcelTable(

    data=serialized_data,
    
    columns=dict(
        oscar='alpha.oscar',
        state_name='alpha.papa.romeo.name',
        statehood_granted='alpha.papa.romeo.statehood_granted',
        state_bird='alpha.papa.romeo.symbols.bird',
        state_flower='alpha.papa.romeo.symbols.flower',
        other_states_bird='alpha.papa.sierra.symbols.bird',
        bravo=None,
        charlie=None,
        average_bravo_charlie=dict(
            header='Avg of Bravo/Charlie',
            data_accessor=lambda item: None,
            formula='=AVERAGE({bravo}, {charlie})',
        ),
    )
)
```

### Co-location of Column and Data Accessor Code

One advantage to this approach is that everything pertaining to a single column is **in one place**! The alternative approach force columns to be specified in one place, and data generated in another. This means that to make a change to a column or columns, you have to change it in multiple places. As the size of your tables grow, it becomes more difficult to maintain.

With this style, columns are defined with the logic used to generate the data in each row _for that column_. If columns need to be reordered, headers renamed, or additional info updated, there is only one place that these changes need to be made.

### Easily Flatten Nested Data Into Row

The next advantage is that it provides a concise, readable style for accessing nested data by chaining properties together (dot-syntax is the default, but custom separators can be specified). This makes it much easier to spot inconsistencies across multiple columns, and identify why a cell is showing up blank or generating an AttributeError.

### Unbreakable Formula References

With XlsxWriter, column formulas can be defined with references to other columns in the table (`formula='=SUM([@Qty] * [@Price]'`). Hard-coding column headers is a bad idea, however, as this is subject to break if the referenced column header changes. In this case, it can be difficult to notice that a formula broke, especially in large Excel tables.

This package solves this by using the columns' keys to generate dynamic references to each column in column formulas. This approach will fail at runtime if the keys change, alerting you that a change needs to be made.

## Nesting Classes and Dicts

This package is flexible enough to handle both class instances and dicts. Classes can be nested inside of dicts (`romeo` is an instance of `State` in the example). Likewise, dicts can be properties of class instances (`symbols`, a `dict`, is a property of each `State` instance). The same syntax is used to access nested values of both classes and dicts.

## Field Separator Syntax

The default field separator is the dot (`.`), but custom characters can be specified. For instance, to assimilate Django's ORM-style "dunder" syntax for querying fields, use `separator='__'`. Columns would then use this like so:

``` py
...
separator='__',
columns=dict(
    oscar='alpha__oscar',
    state_name='alpha__papa__romeo__name',
    ...
```

## Specifying `column=None`

In the simplest cases where the column key is also the top-level attribute that is desired, and the column key is the desired column header, set the column's value to `None`.

For example, these column definitions will generate the header, 'Bravo', and access data on each item using `item.bravo` (class) or `item['bravo']` (dict):

``` py
bravo=None,
bravo={},
bravo=(),
```

## AttributeErrors and KeyErrors
By default, if attributes or keys cannot be found, they fail gracefully - meaning they return the value `None`, and cell values for those fields are blank. This is useful for cases where the shape of each item _is not expected_ to conform perfectly to the column schema.

For debugging purposes, or in other cases where the data _is expected_ to be uniform, you can set `raise_attribute_errors=True`:

``` py
excel_table = ExcelTable(
    ...
    raise_attribute_errors=True,
    
    columns=dict(
        alpha_quebec='alpha.quebec',
        ...
    )
)
```

This will result in:

![image](https://user-images.githubusercontent.com/36553266/148541356-d94f8a70-d972-46db-bea5-296539b791bb.png)



## Other Exceptions
Any other error is printed to the cell in which it occurred, to help diagnose.

## Column Header Text

The `header` attribute defaults to the title-cased dictionary key, unless a header is explicitly provided. For example:

``` py
THIS COLUMN                                 BECOMES THIS HEADER
-------------------------------------------------------------------
oscar=...,                              --> 'Oscar'
state_name=...,                         --> 'State Name'

average_bravo_charlie=dict(             --> 'Avg of Bravo/Charlie'
    header='Avg of Bravo/Charlie',
    ...
)
```

## Additional Column Attributes

Column attributes can be supplied in each column's dictionary, following XlsxWriter's docs. With the exception of `formula`, these attributes simply get passed through to XlsxWriter.

## Column Formulas

Formulas can be specified per XlsxWriter's docs. To dynamically reference the calculated column header of another column in a formula, use curly braces and the column's kwarg.

For instance, the following code for `average_bravo_charlie`...

``` py
bravo=None,
charlie=None,
average_bravo_charlie=dict(
    ...
    formula='=AVERAGE({bravo}, {charlie})',
),
```

...will generate this column formula:

``` py
'=AVERAGE([@[Bravo]], [@[Charlie]])
```

This means that changing the header text in a referenced column will _not_ break the formula! Further, changing the column's kwarg _will_ break the formula if it is not also updated. However, it will raise an error at runtime, rather than failing silently in the Excel file.

## Saving to Excel

To save the data to an Excel file, use XlsxWriter's `worksheet.add_table()` method as usual.

The ExcelTable class automatically calculates the top, left, bottom, and right coordinates of the table based on the size of the data. These values are available in `excel_table.coordinates` as a tuple, making it easy to spread them into the `add_table()` call.

``` py
import xlsxwriter

workbook = xlsxwriter.Workbook('example.xlsx')

worksheet = workbook.add_worksheet()

excel_table = ExcelTable(...)

worksheet.add_table(
    *excel_table.coordinates,
    {
        'columns': excel_table.columns,
        'data': excel_table.data,
        'total_row': excel_table.include_total_row,
        ...
    }
)

workbook.close()
```
