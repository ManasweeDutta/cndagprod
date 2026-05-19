import logging
import re
from pathlib import Path
from typing import Optional, Dict, List

import docx
import docx2txt
from docx.table import Table
from docx.text.paragraph import Paragraph


class DocxReaderAndParserSectionTree:
    """Docx parser."""

    def __init__(self):
        self.section_name = "NA"
        self.jinja_tag = "NA"
        self.index_flag = False

    def load_data(self, file: Path, extra_info: Optional[Dict] = None) -> List[Dict]:
        """Parse file."""
        if extra_info is None:
            extra_info = {}
        docs = self.parse_document(file, extra_info)
        return self.remove_duplicates_based_on_criteria(docs)

    def parse_document(self, file: Path, extra_info: Dict) -> List[Dict]:
        documents = []
        try:
            document = docx.Document(file)
            text = docx2txt.process(file)
            current_levels = [0] * 10
            count = 0

            for block in self.iter_block_items(document):
                if isinstance(block, Paragraph):
                    self.process_paragraph(block, current_levels, extra_info, documents, text, count)
                elif isinstance(block, Table):
                    self.process_table(block, current_levels, extra_info, documents)
                count += 1

            self.add_document_info(documents, extra_info, current_levels)
        except Exception as e:
            logging.error(f"Error parsing document: {e}")
        return documents

    def iter_block_items(self, parent):
        from docx.oxml.table import CT_Tbl
        from docx.oxml.text.paragraph import CT_P

        parent_elm = self.get_parent_element(parent)
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    def get_parent_element(self, parent):
        from docx.document import Document
        from docx.table import _Cell, _Row
        if isinstance(parent, Document):
            return parent.element.body
        elif isinstance(parent, _Cell):
            return parent._tc
        elif isinstance(parent, _Row):
            return parent._tr
        else:
            raise ValueError("Unrecognized parent type")

    def process_paragraph(self, block, current_levels, extra_info, documents,
                          text, count):
        level_from_style_name = {f'Heading {i}': i for i in range(10)}
        if block.style.name not in level_from_style_name and block.style.name != "Heading":
            if "{{" in block.text:
                self.jinja_tag = self.extract_text_from_jinja_tag(block.text)
        else:
            if count > 0:
                self.add_document_info(documents, extra_info, current_levels)
            level = 0 if block.style.name == "Heading" else level_from_style_name[block.style.name]
            if not self.index_flag:
                self.index_flag = self.update_index_flag(text, block.text, level, current_levels)
            if self.index_flag:
                current_levels[level] += 1
                self.reset_lower_levels(current_levels, level)
            self.section_name = block.text.strip()
            self.jinja_tag = "NA"

    def process_table(self, block, current_levels, extra_info, documents):
        header, tmpHeader, value_paragraph, table_mapping, table_header, table_flag, paragraph_flag, inset_table_flag, skipheaderrowcount = [], {}, [], {}, "", False, False, False, 0
        section_index = self.format_levels(current_levels)
        count, table_type = 0, "table"
        header, data_start_row = self.extract_headers_from_table(block)
        for row in block.rows[data_start_row:]:
            row_data, colindex = [], 0
            for cell in row.cells:
                cell_text = ""
                for paragraph in cell.paragraphs:
                    cell_text = cell_text + paragraph.text + " "
                if "%tr endfor" in cell_text:
                    table_flag = False
                    inset_table_flag = True
                elif "%tr for" in cell_text:
                    table_flag, table_type = True, "table"
                    table_header = cell_text.split()[4]
                elif table_flag:
                    if "loop.index" not in cell_text:
                        tmp_jinja_tag = self.extract_text_from_jinja_tag(cell_text).split(".")[-1]
                        table_mapping = f"{table_header}.{tmp_jinja_tag}"
                        self.add_table_document_info(documents, extra_info, header, section_index, table_mapping,
                                                     table_type,
                                                     value_paragraph, colindex)
                    colindex += 1
                    table_mapping = {}
                elif paragraph_flag:
                    row_data.append(cell_text.strip())
                    if "{{" in cell_text:
                        table_type = "text"
                        tmp_jinja_tag = self.extract_text_from_jinja_tag(cell_text)
                        table_mapping = tmp_jinja_tag
                        value_paragraph = row_data[0]
                        self.add_table_document_info(documents, extra_info, header, section_index, table_mapping,
                                                     table_type,
                                                     value_paragraph, colindex)
                        table_mapping = {}
                    colindex += 1
                else:
                    row_data.append(cell_text)
                    paragraph_flag = True
                    colindex += 1

    def add_document_info(self, documents, extra_info, current_levels):
        section_index = self.format_levels(current_levels)
        document_info = extra_info.copy()
        document_info.update({
            "fields": "NA",
            "sectionName": self.section_name,
            "sectionNumber": section_index,
            "jinjaTag": self.jinja_tag,
            "contentType": "text",
            "sectionType": "paragraph"
        })
        documents.append(document_info)

    def add_table_document_info(self, documents, extra_info, header, section_index, table_mapping, table_type,
                                value_paragraph, colIndex=0):
        if table_type == "text" and not value_paragraph:
            header, table_mapping = "NA", "NA"
        elif table_type == "table" and not table_mapping:
            header, table_mapping = "NA", "NA"
        document_info = extra_info.copy()
        document_info.update({
            "sectionName": self.section_name,
            "sectionNumber": section_index,
            "tableHeader": header,
            "jinjaTag": table_mapping,
            "contentType": table_type,
            "sectionType": "table"
        })

        if table_type == "text":
            document_info.update(
                {"fields": value_paragraph.strip(), "fieldLabel": header[0], "columnLabel": header[colIndex]})
        else:
            document_info.update({"fields": header[colIndex]})
        documents.append(document_info)

    def format_levels(self, cur_lev):
        return '.'.join(str(l) for l in cur_lev if l != 0)

    def update_index_flag(self, text, keyword, headinglevel, current_levels):
        section_index = self.get_section_index(text, keyword, headinglevel, current_levels[headinglevel])
        return bool(re.search(r'\d', section_index))

    def reset_lower_levels(self, current_levels, level):
        for l in range(level + 1, 10):
            current_levels[l] = 0

    def get_section_index(self, text, keyword, heading_level, prev_num):
        index = text.find(keyword.strip())
        count = 0
        while index != -1:
            section_index = text[max(0, index - 10):index + len(keyword)].split('\n\n')[-1].strip('\t').split('\t')[0]
            tmp = section_index.split(".")
            if len(tmp) >= heading_level or count >= 3:
                break
            index = text.find(keyword.strip(), index + len(keyword))
            count += 1
        else:
            section_index = ""
        return section_index

    def extract_text_from_jinja_tag(self, text):
        start_index = text.find("{{")
        if start_index == -1:
            return "NA"
        end_index = text.find("}}", start_index + 2)
        if end_index == -1:
            return "NA"
        return text[start_index + 2:end_index].strip().split("|")[0]

    def is_vertically_merged(self, cell):
        """Check if the cell is part of a vertical merge."""
        v_merge_elements = cell._element.xpath('.//w:vMerge')
        if v_merge_elements:
            # Check if it's part of a continued vertical merge (val="continue") or the start of a merge
            return v_merge_elements[0].get("w:val") != "restart"
        return False

    def find_last_header_row(self, table):
        """Find the last header row when vertical merges exist."""
        header_row_index = 0
        num_columns = len(table.rows[0].cells)
        merged_flags = [False] * num_columns
        mergeFlag = False

        for row_index, row in enumerate(table.rows):
            all_columns_finalized = True  # Assume we are at the last row of headers

            for col_index, cell in enumerate(row.cells):
                # Update merge flags for each column
                if self.is_vertically_merged(cell):
                    merged_flags[col_index] = True
                    mergeFlag = True
                else:
                    merged_flags[col_index] = False

                # If any column is still vertically merged, it's not the last header row
                if merged_flags[col_index]:
                    all_columns_finalized = False

            # If all columns are unmerged, we've found the last header row
            if all_columns_finalized and not mergeFlag:
                header_row_index = row_index
                break
            elif all_columns_finalized and mergeFlag:
                header_row_index = row_index - 1
                break

        return header_row_index

    def extract_headers_from_table(self, table):
        """Extract headers based on whether columns are vertically merged or not."""
        headers = []
        """Extract headers based on vertical merge logic."""
        # Find the last header row when columns are vertically merged
        header_row_index = self.find_last_header_row(table)

        # Extract headers from the identified row
        headers = [cell.text.strip() for cell in table.rows[header_row_index].cells]

        # Identify where data rows start (first row after the header)
        data_start_row = header_row_index + 1

        return headers, data_start_row

    def remove_duplicates_based_on_criteria(self, data):
        # Create a set to track unique (sectionName, sectionNumber) pairs with the specified "field_jinja_mapping"
        unique_entries = set()
        filtered_data = []

        for entry in data:
            key = (entry['sectionName'], entry['sectionNumber'])
            if entry['jinjaTag'] == "NA":
                if key not in unique_entries:
                    filtered_data.append(entry)
                    unique_entries.add(key)
            else:
                filtered_data.append(entry)

        return filtered_data
