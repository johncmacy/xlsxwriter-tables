# xlsxwriter-tables
Easily export nested data to a 2-dimensional Excel table

This class is intended to be used with [XlsxWriter](https://xlsxwriter.readthedocs.io/working_with_tables.html). It serves two purposes:

1. Co-location of column info and data-generation logic
2. Easily specify deeply nested data as the source for column data

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

## Nesting Classes and Dicts

The class is flexible enough to handle both class instances and dicts. Classes can be nested inside of dicts (`romeo` is an instance of `State` in the example). Likewise, dicts can be properties of class instances (`symbols`, a `dict`, is a property of each `State` instance). The same syntax is used to access nested values of both classes and dicts.

## Nested Fields Syntax

The default separator character is the dot (`.`). Custom characters can be specified. For instance, to assimilate Django's ORM-style "dunder" syntax for querying fields, use `separator='__'`. Columns would then use this like so:

``` py
...
separator='__',
columns=dict(
    oscar='alpha__oscar',
    state_name='alpha__papa__romeo__name',
    ...
```

## Attribute and Key Errors
If attributes or keys cannot be found, they fail gracefully - meaning they return the value `None`, and cell values for those fields are blank.

Any other error is printed to the cell in which it occurred, to help diagnose.

## Column Header Text

The `header` attribute defaults to the title-cased dictionary key, unless a header is explicitly provided. For example:

``` py
THIS COLUMN                                 BECOMES THIS VALUE
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

