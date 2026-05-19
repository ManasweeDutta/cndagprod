import logging
import os
import re
from collections import defaultdict
from pathlib import Path
from typing import Any

import docx
import docx2txt
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.table import Table
from docx.text.paragraph import Paragraph
from eric_sdui_common_logger_py import setup_logging


class JinjaTagAddInDocx:
    """Docx parser."""

    def __init__(self):
        self.section_name = "NA"
        self.jinja_tag = "NA"
        self.index_flag = False
        self.add_jinja_paragraph_flag = False
        self.document = None
        # Define a regex pattern to match any character that is not alphanumeric
        self.special_char_pattern = re.compile(r'[\W\u00A0]+')
        self.table_counter = defaultdict(int)

    def load_data(self, file: Path) -> tuple[Any, Any]:
        """Parse file."""
        jinja_file_name, jinja_doc_path = self.parse_document(file)
        return jinja_file_name, jinja_doc_path

    def parse_document(self, file: Path) -> tuple[None, str] | tuple[str, str]:
        try:
            file_name = file.stem
            self.document = docx.Document(file)
            text = docx2txt.process(file)
            current_levels = [0] * 10
            for block in self.iter_block_items(self.document):
                if isinstance(block, Paragraph):
                    self.process_paragraph(block, current_levels, text)
                elif isinstance(block, Table):
                    self.process_table(block, current_levels)
        except Exception as e:
            logging.error(f"Error parsing base line document: {e}")
            return None, "Error"
        jinja_doc_path: str = '/'.join([os.path.dirname(file), file_name + "_with_jinja.docx"])
        jinja_file_name = file_name + "_with_jinja.docx"
        self.document.save(jinja_doc_path)
        return jinja_file_name, jinja_doc_path

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

    # this is function to use to append jinja tag on section header/sub-header (Header 1, Header 2...so on) in form of section name + Underscore+ section index + _text/escape, if section index is None or empty not added into jinja_tag
    def process_paragraph(self, block, current_levels, text):
        level_from_style_name = {f'Heading {i}': i for i in range(10)}
        if block.style.name in level_from_style_name or block.style.name == "Heading":
            level = 0 if block.style.name == "Heading" else level_from_style_name[block.style.name]
            if not self.index_flag:
                self.index_flag = self.update_index_flag(text, block.text, level, current_levels)
            if self.index_flag:
                current_levels[level] += 1
                self.reset_lower_levels(current_levels, level)
            self.section_name = block.text.strip()
            secName = re.sub(self.special_char_pattern, '_', self.section_name)
            secName = self.format_jinja_tag(secName)
            section_index = self.format_levels(current_levels)
            secIndex = re.sub(self.special_char_pattern, '_', section_index)
            jinja_tag = f"\n{{{{ pg_{secName}_{secIndex}_text|escape }}}}"
            jinja_tag = self.remove_extra_underscores(jinja_tag)
            self.append_jinja_to_paragraph(block, jinja_tag)

    # this is function to use to append jinja tag in tabular content on section header/sub-header (Header 1, Header 2...so on) in form of section name + Underscore+ section index + table header+ _table/escape, if section index is None or empty not added into jinja_tag
    def process_table(self, block, current_levels):
        section_index = None
        try:
            section_index = self.format_levels(current_levels)
            if self.table_counter.get(section_index) is not None:
                self.table_counter[section_index] += 1
            else:
                self.table_counter[section_index] = 1
            count = 0
            table_cell_flag = False
            rowcount, colcount = 0, 0
            header, data_start_row = self.extract_headers_from_table(block)
            if data_start_row > 1:
                count = data_start_row - 1
            else:
                count = data_start_row

            # check for table cell
            table_cell_flag = self.is_table_cell_type(block)
            if table_cell_flag:
                self.table_cell_jinja_generate(block, data_start_row, section_index, header)
            else:
                if len(block.rows) - data_start_row < 3:
                    last_row = block.rows[-1]
                    for _ in range(len(block.rows) - data_start_row + 1):
                        new_row = block.add_row()
                        for cell_index, cell in enumerate(last_row.cells):
                            new_row.cells[cell_index].text = ""
                            # new_row.cells[cell_index].paragraphs[0].style = cell.paragraphs[0].style
                            new_row.cells[cell_index].width = docx.shared.Inches(2)
                            # new_row.cells[cell_index].height = docx.shared.Inches(1)
                        # Apply height to the new row
                        new_row.height = docx.shared.Inches(0.5)
                        new_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

            for row in block.rows[data_start_row:]:
                row_data, colindex = [], 0
                for cell in row.cells:
                    cell_text = ""
                    for paragraph in cell.paragraphs:
                        cell_text = cell_text + paragraph.text + " "
                    if count > 0:
                        if cell_text.strip():
                            table_cell_flag = True
                        elif not cell_text.strip() and not table_cell_flag:
                            if count == 1:
                                row.cells[0].merge(row.cells[colindex])
                                secName = re.sub(self.special_char_pattern, '_', self.section_name)
                                secName = self.format_jinja_tag(secName)
                                secIndex = re.sub(self.special_char_pattern, '_', section_index)
                                tmptext = f"%tr for tmp in tb_{secName}_{secIndex}_table{self.table_counter[section_index]} %"
                                tmptext = "{" + tmptext + "}"
                                tmptext = self.remove_extra_underscores(tmptext)
                                paragraph = cell.paragraphs[0]
                                run = paragraph.add_run(tmptext)
                                run.font.size = Pt(10)
                            # this condition will execute if table has header but None header column contains value in rows. This will be executed in 3rd row after table header reading and append jinja %tr enffor
                            elif count == 3:
                                row.cells[0].merge(row.cells[colindex])
                                tmptext = "%tr endfor %"
                                tmptext = "{" + tmptext + "}"
                                paragraph = cell.paragraphs[0]
                                run = paragraph.add_run(tmptext)
                                run.font.size = Pt(10)
                            # this condition will execute if table has header but None header column contains value in rows. This will be executed in 2rd row after table header reading and append jinja {{ tmp.header name}}. header name must replace with underscore if any special characters
                            elif count == 2:
                                tmpheader = self.format_jinja_tag(header[colindex])
                                if header[colindex] in ["#"]:
                                    tmptext = "loop.index"
                                else:
                                    tmptext = f"tmp.tb_{tmpheader}"
                                tmptext = "{{ " + tmptext + "|escape }}"
                                tmptext = self.remove_extra_underscores(tmptext)
                                paragraph = cell.paragraphs[0]
                                run = paragraph.add_run(tmptext)
                                run.font.size = Pt(10)
                            # this condition will execute if table has header but None header column contains value in rows. row will be removed after 3rd row of table header reading.
                            elif count > 3:
                                if row._element.getparent():
                                    row._element.getparent().remove(row._element)
                        row_data.append(cell_text.strip())
                        colindex += 1
                rowcount = rowcount + 1
                count = count + 1
        except Exception as e:
            logging.error(f"Error generating jinja for index [{section_index}]: {e}")

    def table_cell_jinja_generate(self, block, data_start_row, section_index, header):
        """
        Processes rows in a table block, starting from a specific row, and updates empty cells with Jinja2 placeholders.

        Args:
            block: The table block containing rows and cells.
            data_start_row: The starting row index for processing.
            section_index: The section index used in the Jinja2 placeholder.
            header: The list of column headers.
            special_char_pattern: Regex pattern to replace special characters.
        """
        for row in block.rows[data_start_row:]:
            row_data, col_index = [], 0

            for cell in row.cells:
                cell_text = " ".join(paragraph.text for paragraph in cell.paragraphs).strip()

                if not cell_text:
                    sec_name = (
                        re.sub(self.special_char_pattern, '_', row_data[0]) if row_data else ""
                    )
                    sec_name = self.format_jinja_tag(sec_name).replace(" ", "_")

                    sec_index = re.sub(self.special_char_pattern, '_', section_index)
                    tmp_header = self.format_jinja_tag(header[col_index])

                    # Construct the Jinja2 placeholder
                    tmp_text = f"tc_{sec_name}_{sec_index}_{tmp_header}_text|escape"
                    tmp_text = f"{{{{ {tmp_text} }}}}"
                    tmp_text = self.remove_extra_underscores(tmp_text)

                    # Add the placeholder to the cell
                    paragraph = cell.paragraphs[0]
                    run = paragraph.add_run(tmp_text)
                    run.font.size = Pt(10)  # Set font size for the placeholder

                # Append the cell text (or placeholder) to the row data
                row_data.append(cell_text)
                col_index += 1

    def is_table_cell_type(self, table):
        """
        Determines if a table contains any data in cells after the header row.

        Args:
            table: A docx Table object to check

        Returns:
            bool: True if any cell after the header row contains data, False otherwise

        This method checks if a table has any non-empty cells after the header row.
        Used to determine if table should be processed as a data table vs an empty template.
        Skips the first row (header) and returns True as soon as it finds any cell with content.
        """
        # Iterate through the rows of the table, starting from the second row
        for i, row in enumerate(table.rows):
            if i == 0:  # Skip the header row
                continue

            # Check if any cell in the row has data
            if any(cell.text.strip() for cell in row.cells):
                return True

        # If no data is found in any cell after the header row
        return False

    def format_jinja_tag(self, text):
        # Remove special characters and prevent double underscores
        text = re.sub(self.special_char_pattern, '_', text).replace(" ", "_").replace("__", "_")
        text = re.sub(r'\n', '_', text)
        return text.strip('_')  # Remove leading/trailing underscores

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

    def append_jinja_to_paragraph(self, paragraph, jinja_tag):
        new_paragraph = OxmlElement('w:p')
        paragraph._element.addnext(new_paragraph)
        paragraph = docx.text.paragraph.Paragraph(new_paragraph, self.document)
        run = paragraph.add_run(f"{jinja_tag} \n")
        run.font.size = Pt(10)  # Set font size for Jinja tags
        run.bold = False
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT  # Set left alignment

    def format_levels(self, cur_lev):
        return '.'.join(str(l) for l in cur_lev if l != 0)

    def remove_extra_underscores(self, text):
        # Replace multiple underscores with a single underscore
        return re.sub(r'__+', '_', text)

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