import logging
from typing import Optional

import openpyxl
import pandas as pd
from eric_sdui_docgen_libcommon_py.utils import (
    convert_mixed_types_to_string,
    find_and_convert_invalid_datetime_to_time,
    clean_and_remove_empty_rows,
)


class ExcelParser:

    def __new__(cls, file_path, read_only_mode=True, data_only_mode=True):
        instance = super(ExcelParser, cls).__new__(cls)
        try:
            instance.file_path = file_path
            instance.workbook = openpyxl.load_workbook(
                file_path, data_only=data_only_mode, read_only=read_only_mode
            )
            # Access the document properties
            properties = instance.workbook.properties

            instance.metadata_info = {
                "created_by": properties.creator,
                "updated_by": properties.lastModifiedBy,
                "created_on": properties.created.strftime("%Y-%m-%d %H:%M:%S"),
                "updated_on": properties.modified.strftime("%Y-%m-%d %H:%M:%S"),
            }

        except Exception as e:
            logging.error(f"Failed to load workbook: {e}")
            return None
        return instance

    def get_data_from_defined_name(self, defined_name) -> Optional[pd.DataFrame]:

        # Check if the named range exists
        if defined_name not in self.workbook.defined_names:
            logging.error(
                f"Defined named range '{defined_name}' not found in the workbook."
            )
            return None

        # Get the named range
        named_range = self.workbook.defined_names[defined_name]

        # Initialize list to hold data from the named range
        range_data = []

        # Iterate over the destinations (can be multiple ranges)
        for title, coord in named_range.destinations:
            sheet = self.workbook[title]
            for row in sheet[coord]:
                range_data.append([cell.value for cell in row])

        if not range_data:
            logging.error(f"No data found in the named range '{defined_name}'.")
            return None

        # Create a DataFrame using the first row as the header
        data = clean_and_remove_empty_rows(range_data[1:])
        data = [[""] * len(range_data[0])] if not data else data
        df = pd.DataFrame(data, columns=range_data[0])
        # fillna not working with datetime, existing pandas bug 'https://github.com/pandas-dev/pandas/issues/11953'
        df = df.fillna(" ")
        df = df.replace(" ", "")

        return df

    def get_key_value_data_from_defined_name(
        self, defined_name, key_field_ref, value_field_ref
    ) -> Optional[dict]:
        data_df = self.get_data_from_defined_name(defined_name)
        if data_df is not None:
            data_dict = dict(zip(data_df[key_field_ref], data_df[value_field_ref]))

            strip_key_dict = {}
            for key, value in data_dict.items():
                if isinstance(key, str):
                    strip_key_dict[key.strip()] = value
                else:
                    strip_key_dict[str(key)] = value

            return strip_key_dict
        return None

    def get_json_array_data_from_defined_name(self, defined_name):
        data_df = self.get_data_from_defined_name(defined_name)
        if data_df is not None:
            data_df.columns = data_df.columns.str.strip()
            data_df = find_and_convert_invalid_datetime_to_time(data_df)
            data_dict = data_df.to_dict(orient="records")
            data_dict = convert_mixed_types_to_string(data_dict)
            return data_dict
        return None
