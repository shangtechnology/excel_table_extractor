"""
MIT License

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Author: James Shang
"""
import pandas as pd
from typing import List, Tuple, Optional


class ExcelTableExtractor:
    def __init__(self, file_path: str, keywords: Optional[List[str]] = None) -> None:
        """
        Initializes the ExcelTableExtractor with the given file path and optional keywords.

        :param file_path: Path to the Excel file to be processed.
        :param keywords: Optional list of keywords to help identify header rows.
        """
        self.file_path = file_path
        self.keywords = keywords if keywords is not None else ['date', 'name', 'amount', 'total', 'price']
        self.data_frame = pd.read_excel(self.file_path, header=None)

    def is_header_row(self, row: pd.Series, next_row: pd.Series) -> bool:
        """
        Determines if a row is a header by checking if it contains primarily strings or date-like values,
        or specific keywords, and if the next row contains numeric data indicating a data row.

        :param row: A row from the DataFrame.
        :param next_row: The next row from the DataFrame.
        :return: True if the row is likely a header, False otherwise.
        """
        threshold = 0.75
        non_null_values = row.dropna()
        if len(non_null_values) == 0:
            return False

        # Check if values are strings or can be converted to dates
        def is_string_or_date(value):
            if isinstance(value, str):
                return True
            try:
                pd.to_datetime(value, errors='raise')
                return True
            except (ValueError, TypeError):
                return False

        string_or_date_count = non_null_values.apply(is_string_or_date).sum()

        # Check if any keyword is present in the row
        keyword_present = any(
            keyword.lower() in str(value).lower() for value in non_null_values for keyword in self.keywords)

        # Check if the next row is likely a data row
        next_row_non_null = next_row.dropna()
        if len(next_row_non_null) == 0:
            return False

        next_row_numeric_count = next_row_non_null.apply(lambda x: isinstance(x, (int, float))).sum()
        next_row_is_data = next_row_numeric_count / len(next_row_non_null) >= threshold

        return (string_or_date_count / len(non_null_values) >= threshold or keyword_present) and next_row_is_data

    def detect_tables(self) -> Tuple[List[pd.DataFrame], List[int], List[List[str]]]:
        """
        Detects tables within the Excel file by identifying headers and subsequent data rows.

        :return: A tuple containing a list of DataFrames representing tables, a list of header row indices, and a list of headers.
        """
        tables = []
        current_table = []
        header_row_indices = []
        headers = []
        header = None
        in_table = False

        for index in range(len(self.data_frame)):
            row = self.data_frame.iloc[index]
            next_row = self.data_frame.iloc[index + 1] if index + 1 < len(self.data_frame) else pd.Series(
                [None] * len(row))
            print(f"Checking row {index}: {row.tolist()}")
            if self.is_header_row(row, next_row) and not in_table:  # Ensure we are not already in a table
                print(f"Header detected at row {index}")
                if current_table:
                    tables.append(pd.DataFrame(current_table, columns=header))
                    current_table = []
                header_row_indices.append(index)
                header = row.dropna().tolist()
                headers.append(header)
                in_table = True
            elif in_table and len(row.dropna()) == len(header):
                # Align the data row correctly by considering the header length
                current_table.append(row.dropna().tolist())  # Drop NaN and keep the rest
            else:
                if current_table:
                    tables.append(pd.DataFrame(current_table, columns=header))
                    current_table = []
                header = None
                in_table = False

        if current_table:
            tables.append(pd.DataFrame(current_table, columns=header))

        return tables, header_row_indices, headers

    def process_tables(self, tables: List[pd.DataFrame], headers: List[List[str]]) -> List[pd.DataFrame]:
        """
        Processes detected tables to assign column headers and clean the data.

        :param tables: A list of DataFrames representing detected tables.
        :param headers: A list of headers for the tables.
        :return: A list of processed DataFrames with headers assigned.
        """
        processed_tables = []
        for table, header in zip(tables, headers):
            print(f"Assigning headers: {header}")
            if len(header) == table.shape[1]:
                table.columns = header
            else:
                print(f"Header length {len(header)} does not match table width {table.shape[1]}")
            processed_tables.append(table.reset_index(drop=True))
        return processed_tables

    def extract_tables(self) -> List[pd.DataFrame]:
        """
        Extracts tables from the Excel file, processing and returning them as a list of DataFrames.

        :return: A list of DataFrames representing the extracted tables.
        """
        tables, header_row_indices, headers = self.detect_tables()
        processed_tables = self.process_tables(tables, headers)
        return processed_tables


# Example usage
if __name__ == "__main__":
    file_path = "multi_table_file.xlsx"
    extractor = ExcelTableExtractor(file_path)
    tables = extractor.extract_tables()

    for i, table in enumerate(tables):
        print(f"Table {i + 1}:")
        print(table)
        print("\n")
