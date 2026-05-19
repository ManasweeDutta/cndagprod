import json
import logging
from datetime import datetime, time, date
import pandas as pd


def clean_and_remove_empty_rows(data):
    cleaned_data = []
    for row in data:
        cleaned_row = ["" if element is None else element for element in row]
        # Check if the row contains all empty strings
        if any(element != "" for element in cleaned_row):
            cleaned_data.append(cleaned_row)
    return cleaned_data


class JsonTimeDataEncoder(json.JSONEncoder):
    @staticmethod
    def format_time(obj, default=str):
        if isinstance(obj, pd.Timestamp):
            return obj.strftime('%Y-%m-%d')
        elif isinstance(obj, datetime):
            return obj.strftime('%Y-%m-%d')  # Serialize datetime to custom format
            # return obj.strftime('%Y-%m-%d %H:%M:%S')  # Serialize datetime to custom format
        elif isinstance(obj, time):
            return obj.strftime('%H:%M')  # Serialize time to HH:MM:SS format
        else:
            return default(obj)

    def default(self, obj):
        return self.format_time(obj, super().default)


# Function to convert datetime.datetime to datetime.time if the date is 1900-01-01
def convert_invalid_datetime_to_time(value):
    if isinstance(value, datetime) and value.date() == date(1900, 1, 1):
        return value.time()
    return value


# Function to find the column with mixed types and convert the values
def find_and_convert_invalid_datetime_to_time(df):
    for column in df.columns:
        if any(isinstance(val, datetime) and val.date() == date(1900, 1, 1) for val in df[column]):
            df[column] = df[column].apply(convert_invalid_datetime_to_time)
            logging.info(f"column {column} having invalid datetime object, convert it to time")
    return df


def convert_mixed_types_to_string(json_array):
    # Function to check if the values in a list have mixed types
    def has_mixed_types(values):
        types = set(type(value) for value in values if value != "")
        return len(types) > 1

    # Detect keys with mixed types
    keys_with_mixed_types = set()
    for item in json_array:
        for key, value in item.items():
            if key not in keys_with_mixed_types:
                values = [i.get(key) for i in json_array if i.get(key) is not None]
                if has_mixed_types(values):
                    keys_with_mixed_types.add(key)

    # Convert the values of the detected keys to text
    for item in json_array:
        for key in keys_with_mixed_types:
            if key in item:
                item[key] = JsonTimeDataEncoder.format_time(item[key])

    return json_array
