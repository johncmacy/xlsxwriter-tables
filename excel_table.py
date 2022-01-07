from typing import Union

class ExcelTable:
    def _get_column(self, column_name:str, column_props:Union[dict, str]) -> dict:

        '''
        Defaults to the title-cased `column_name`:

            my_favorite_color: 'My Favorite Color'

        For acronyms, abbreviations, provide the correct capitalization in the `column_name`:

            NFL_franchise_name: 'NFL Franchise Name'

        '''

        column = {'header': column_name.replace('_', ' ').title()}

        '''
        If other attributes were provided, such as `formula`, or `format`, pass them along.
        '''

        if isinstance(column_props, dict):
            column.update(**column_props)

        return column

    def _get_data(self, item, column_name:str, column_props:Union[None, dict, str, tuple], separator, raise_attribute_errors):

        '''
        Set the default value to `column_name`:
        
        columns = {
            'date_is_tuesday': ...
        }
        '''

        data_accessor = column_name

        '''
        For each column key, its value `column_props` could be None, or a string, tuple, function, or dictionary:

        columns = {
            'column_with_none': None,

            'column_with_string': 'deeply.nested.property',

            'column_with_tuple': ('deeply', 'nested', 'property'),

            'column_with_function': lambda item: ...,

            'column_with_dict': {
                'data_accessor': ...,
            },
        }

        If `column_props` is a dictionary, it may have a `data_accessor` key.
        If it does, `data_accessor` could be a string, tuple, or function.
        If not, continue to use the `column_name` as the `data_accessor`.
        '''

        if column_props:
            
            '''
            If `column_props` is a dict, look for a `data_accessor` property.

            columns = {
                'date_is_tuesday': {
                    'data_accessor': ...
                }
            }

            `data_accessor` could be a function, str, or tuple.
            '''

            if isinstance(column_props, dict):
                if 'data_accessor' in column_props.keys():
                    data_accessor = column_props['data_accessor']

                else:
                    '''
                    If `column_props` is a dict, but it doesn't have a
                    `data_accessor` key, then use the `column_name` as
                    a string as the `data_accessor`.
                    '''
                    pass

            else:
                '''
                If not a dict, it's either a string, tuple, or function.
                '''
                
                data_accessor = column_props


        '''
        If `data_accessor` is a function, call the function and
        return the resulting value.

        Note: The function should expect a single kwarg, `item`.
        
        Example:

        def day_of_week_is_tuesday(item):
            return item.start_date.weekday() == 1

        columns = {
            'date_is_tuesday': {
                'data_accessor': day_of_week_is_tuesday,
            }
        }

        Or, as an inline (lambda) function:
        
        columns = {
            'date_is_tuesday': {
                'data_accessor': lambda item: item.start_date.weekday() == 1
            }
        }
        '''

        if callable(data_accessor):
            return data_accessor(item)

        '''
        If we've made it this far, it's either a tuple or a string.
        If it's a string, split it using the separator, and convert to a tuple.

        For the following examples, assume each item has a data structure like so:

            {
                'alpha': {
                    'bravo': {
                        'charlie': 123,
                    }
                }
            }

        The default attribute separator is dot ('.'):

            alpha.bravo.charlie'

        Custom separators can be used. For instance, to resemble Django's ORM, set the separator to '__':

            'alpha__bravo__charlie'

        '''

        if isinstance(data_accessor, str):
            data_accessor = tuple(data_accessor.split(separator))

        '''
        By now, we should have a tuple, which is a list 
        of nested attributes that point to where the data is.

        This code recursively traverses through the tuple of 
        nested attributes and returns the value that is deeply
        nested inside the data structure.
        '''

        if isinstance(data_accessor, tuple):
            # need to deepcopy here?
            nested_data = item

            for key in data_accessor:
                try:
                    if isinstance(nested_data, dict):
                        nested_data = nested_data[key]

                    else:
                        nested_data = getattr(nested_data, key)

                    if callable(nested_data):
                        nested_data = nested_data()

                except (KeyError, AttributeError) as e:
                    if raise_attribute_errors:
                        return f'{type(e)}: {str(e)}'
                    else:
                        return None

                except Exception as e:
                    '''
                    If an exception other than (KeyError, AttributeError) is encountered, the error message
                    is returned and displayed in the cell to aid in troubleshooting.
                    '''
                    
                    return f'{type(e)}: {str(e)}'

            return nested_data

        '''
        If we reach this point, we don't know how to access data from the item, so raise an error.
        '''

        raise ValueError(f'''
            Unable to detect the `data_accessor`. Please provide a function, string, or tuple.
                - column_name={column_name}
                - column_props={column_props}
        ''')

    def __init__(self, columns:dict, data:list, separator='.', include_total_row=True, raise_attribute_errors=False):
        columns_dict = {
            name: self._get_column(name, props)
            for name, props
            in columns.items()
        }

        columns_and_headers = {
            key: f'[@[{value["header"]}]]'
            for key, value
            in columns_dict.items()
        }

        for column in columns_dict.values():
            if 'formula' in column.keys():
                formula_str:str = column['formula']
                column['formula'] = formula_str.format(**columns_and_headers)

        self.columns:list[dict] = tuple(columns_dict.values())

        self.data:list = [
            [
                self._get_data(item, column_name, column_props, separator, raise_attribute_errors) 
                for column_name, column_props 
                in columns.items()
            ] 
            for item 
            in data
        ]
        
        self.top_left = (0,0)
        self.bottom_right = (
            len(self.data) - 1 + 1 + (1 if include_total_row else 0),
            len(self.columns) - 1
        )
        self.coordinates = (*self.top_left, *self.bottom_right)
        
        self.include_total_row = include_total_row
