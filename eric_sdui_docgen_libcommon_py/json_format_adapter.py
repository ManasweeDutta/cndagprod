import logging
import os, re
from collections import defaultdict, OrderedDict
from copy import deepcopy

import pandas as pd
import pypandoc
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from eric_sdui_common_utility_py.utils import sanitize_dict_keys
import markdown
from eric_sdui_common_logger_py import setup_logging
setup_logging("DOC_GENERATION")

reference_docx_path = "/config/reference.docx"

SECTION_CONTAIN_TYPE_MAP = {
    ("table", "text"): ("table_cell", "Text inside a table cell"),
    ("table", "table"): ("table", "Table"),
    ("paragraph", "text"): ("paragraph", "Paragraph"),
}


def get_section_contain_type(new_content_type):
    for key, value in SECTION_CONTAIN_TYPE_MAP.items():
        if value[0] == new_content_type:
            return key
    return None, None


def convert_to_config_template_list(aggregated_config_template):
    # Create a list to hold the reversed data
    config_template_list = []

    # Process each section in the output JSON
    try:
        for section in aggregated_config_template:
            section_name = section["sectionName"]
            section_number = section["sectionNumber"]
            for content in section["content"]:
                new_content_type = content["type"]
                section_type, content_type = get_section_contain_type(new_content_type)
                if section_type is None:
                    continue

                for field in content["fields"]:
                    if new_content_type == "table_cell":
                        new_entry = {
                            "fields": field.pop("name", ""),
                            "tableHeader": content["tableHeader"],
                            "sectionName": section_name,
                            "sectionNumber": section_number,
                            "jinjatag": field.pop("jinjaTag", ""),
                            "contenttype": content_type,
                            "sectionType": section_type,
                            "fieldLabel": field.pop("fieldLabel", ""),
                            "columnLabel": field.pop("columnLabel", ""),
                            "generationType": field.pop("generationType", ""),
                        }
                    elif new_content_type == "table":
                        new_entry = {
                            "fields": field.pop("name", ""),
                            "tableHeader": content["tableHeader"],
                            "sectionName": section_name,
                            "sectionNumber": section_number,
                            "jinjaTag": field.pop("jinjaTag", ""),
                            "contentType": content_type,
                            "sectionType": section_type,
                            "generationType": field.pop("generationType", ""),
                        }
                    else:
                        new_entry = {
                            "fields": field.pop("name", ""),
                            "sectionName": section_name,
                            "sectionNumber": section_number,
                            "jinjaTag": field.pop("jinjaTag", ""),
                            "contentType": content_type,
                            "sectionType": section_type,
                            "generationType": field.pop("generationType", ""),
                        }
                    # Copy any other extra keys in the field
                    new_entry.update(field)

                    config_template_list.append(new_entry)
    except Exception as e:
        logging.error(f"Failed to parse config: {e}")
        raise e

    return config_template_list


def convert_to_aggregate_config_template(config_template_list):
    # Create a dictionary to hold the aggregated data
    aggregated_data = defaultdict(list)
    # Process each entry in the input JSON
    for entry in config_template_list:
        try:
            section_name = entry.pop("sectionName")
            section_number = entry.pop("sectionNumber")
            content_type = entry.pop("contentType")
            section_type = entry.pop("sectionType")
            content_type, label = SECTION_CONTAIN_TYPE_MAP[(section_type, content_type)]
        except Exception as e:
            logging.debug(f"Error: {e}")
            continue

        # Remove invalid jinaja tag
        jinjatag = entry.get("jinjaTag")
        if jinjatag is None or jinjatag == "NA" or not jinjatag:
            continue

        # Create the field entry with remaining keys from the original entry
        field_entry = {
            "name": entry.pop("fields"),
        }

        # Add any remaining keys to the field entry
        field_entry.update(entry)

        # Add Blank functions to the field entry
        if "conditions" not in field_entry:
            field_entry.update({"conditions": [], "generationType": "na",})

        # Append the new entry to the corresponding section and number
        if content_type != "paragraph":
            aggregated_data[(section_number, section_name)].append(
                {
                    "type": content_type,
                    "label": label,
                    "tableHeader": field_entry.pop("tableHeader"),
                    "fields": [field_entry],
                }
            )
        else:
            aggregated_data[(section_number, section_name)].append(
                {"type": content_type, "label": label, "fields": [field_entry]}
            )

    # Merge consecutive entries with the same content_type
    merged_output = []
    for section_info, contents in aggregated_data.items():
        section_number, section_name = section_info
        merged_content = []
        prev_entry = None
        for entry in contents:
            if prev_entry and prev_entry["type"] == entry["type"]:
                prev_entry["fields"].extend(entry["fields"])
            else:
                merged_content.append(entry)
                prev_entry = entry
        merged_output.append(
            {
                "sectionName": section_name,
                "sectionNumber": section_number,
                "content": merged_content
            }
        )

    return merged_output


def split_json_array_by_property(input_json, data_property):
    data_array = []
    metadata_array = []
    if not data_property:
        return input_json, []

    for item in input_json:
        data_entry = {}
        metadata_entry = {}
        for key, value in item.items():
            data_entry[key] = {}
            metadata_entry[key] = {}
            for prop, prop_value in value.items():
                if prop in data_property:
                    data_entry[key][prop] = prop_value
                else:
                    metadata_entry[key][prop] = prop_value
        data_array.append(data_entry)
        metadata_array.append(metadata_entry)
    return data_array, metadata_array


def transform_table_cell(header_list, fields):
    transformed_data_row_entry = OrderedDict()
    # Iterate through each entry in the json_data
    for entry in fields:
        try:
            key = (entry.pop("fieldlabel"), entry.pop("name"))

            value = {
                entry.pop("columnlabel"): {
                    "value": entry.pop("generatedcontent"),
                    "static": False,
                    **entry,
                }
            }
        except Exception as e:
            logging.error(f"Unable to convert the entry {entry}. Error is {str(e)}")

        if key not in transformed_data_row_entry:
            transformed_data_row_entry[key] = value
        else:
            transformed_data_row_entry[key].update(value)

    transformed_data = []

    for key, value in transformed_data_row_entry.items():
        data = OrderedDict({key[0]: {"value": key[1], "static": True}})
        data.update(value)
        missing_col = set(header_list) - set(data.keys())
        for col in missing_col:
            data.update({col: {"value": "", "locked": True}})
        transformed_data.append(data)

    data, metadata = split_json_array_by_property(
        transformed_data, ["value", "static", "locked"]
    )
    return {"generatedcontent": data, "metadata": metadata}


def transform_table(fields):
    metadata = {}
    data_dict = {}
    # check table column is already merged into single entry
    if (
        fields
        and len(fields) == 1
        and isinstance(fields[0], dict)
        and fields[0].get("need_to_merged_table_column", True) is False
    ):
        transformed_data = {
            "generatedcontent": fields[0].pop("generatedcontent"),
            "metadata": fields[0],
        }
        return transformed_data

    for col in fields:
        name = col.pop("name")
        generatedcontent = col.pop("generatedcontent")
        metadata[name] = col
        data_dict[name] = generatedcontent

    # Convert the data dictionary to a pandas DataFrame
    df = pd.DataFrame(data_dict)

    # add cleanup ?
    transformed_data = {
        "metadata": metadata,
        "generatedcontent": df.fillna("").to_dict(orient="records"),
    }
    return transformed_data


def transform_paragraph(fields):
    metadata = {}
    # check table column is already merged into single entry
    if fields and len(fields) == 1 and isinstance(fields[0], dict):
        transformed_data = {
            "generatedcontent": fields[0].pop("generatedcontent"),
            "metadata": fields[0],
        }
        return transformed_data

    # add cleanup ?
    transformed_data = {"metadata": fields, "generatedcontent": ""}
    return transformed_data


def transform_table_and_table_cell_fields(aggregated_config_template):
    transform_data = []
    for section_dict in aggregated_config_template:
        content = section_dict.pop("content")
        modified_content = []
        for item in content:
            if item["type"] == "table_cell":
                tableheader = item["tableheader"]
                fields = item.pop("fields")
                new_item = dict(**item)
                new_item.update(transform_table_cell(tableheader, fields))
                modified_content.append(new_item)
            elif item["type"] == "table":
                fields = item.pop("fields")
                new_item = dict(**item)
                new_item.update(transform_table(fields))
                modified_content.append(new_item)
            elif item["type"] == "paragraph":
                fields = item.pop("fields")
                new_item = dict(**item)
                new_item.update(transform_paragraph(fields))
                modified_content.append(new_item)
            else:
                modified_content.append(item)

        entry = dict(**section_dict)
        entry["content"] = modified_content
        transform_data.append(entry)
    return transform_data


def is_heading(paragraph):
    return paragraph.style.name.startswith("Heading")


def is_numbered(paragraph):
    pPr = paragraph._element.pPr
    if pPr is None:
        return False
    return pPr.find(qn("w:numPr")) is not None

def get_numbering_format(document, numId, ilvl="0"):
    """
    Looks up the numbering format (e.g. 'decimal' or 'bullet') for the given numbering id
    and indentation level (ilvl) in the provided document.
    """
    try:
        numbering = document.part.numbering_part.element
    except (AttributeError, KeyError):
        # Document might not have a numbering part if it's empty or simple
        return None

    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(numId):
            abstractNumId_elem = num.find(qn("w:abstractNumId"))
            if abstractNumId_elem is not None:
                abstractNumId_val = abstractNumId_elem.get(qn("w:val"))
                for abstractNum in numbering.findall(qn("w:abstractNum")):
                    if abstractNum.get(qn("w:abstractNumId")) == abstractNumId_val:
                        # Find the *specific* level (ilvl)
                        for lvl in abstractNum.findall(qn("w:lvl")): # Iterate all levels
                            if lvl.get(qn("w:ilvl")) == ilvl: # Match the level
                                numFmt_elem = lvl.find(qn("w:numFmt"))
                                if numFmt_elem is not None:
                                    return numFmt_elem.get(qn("w:val"))
                        # Fallback: If specific level not found, return first level's format
                        lvl = abstractNum.find(qn("w:lvl"))
                        if lvl is not None:
                            numFmt_elem = lvl.find(qn("w:numFmt"))
                            if numFmt_elem is not None:
                                return numFmt_elem.get(qn("w:val"))
    return None


def get_numbering_format(document, numId, ilvl="0"):
    """
    Looks up the numbering format (e.g. 'decimal' or 'bullet') for the given numbering id
    and indentation level (ilvl) in the provided document.
    """
    try:
        numbering = document.part.numbering_part.element
    except (AttributeError, KeyError):
        # Document might not have a numbering part if it's empty or simple
        logging.warning(f"Document has no numbering part. Cannot find format for numId {numId}.")
        return None

    for num in numbering.findall(qn("w:num")):
        if num.get(qn("w:numId")) == str(numId):
            abstractNumId_elem = num.find(qn("w:abstractNumId"))
            if abstractNumId_elem is not None:
                abstractNumId_val = abstractNumId_elem.get(qn("w:val"))
                for abstractNum in numbering.findall(qn("w:abstractNum")):
                    if abstractNum.get(qn("w:abstractNumId")) == abstractNumId_val:
                        # Find the *specific* level (ilvl)
                        for lvl in abstractNum.findall(qn("w:lvl")): # Iterate all levels
                            if lvl.get(qn("w:ilvl")) == ilvl: # Match the level
                                numFmt_elem = lvl.find(qn("w:numFmt"))
                                if numFmt_elem is not None:
                                    return numFmt_elem.get(qn("w:val"))
                        # Fallback: If specific level not found, return first level's format
                        lvl = abstractNum.find(qn("w:lvl"))
                        if lvl is not None:
                            numFmt_elem = lvl.find(qn("w:numFmt"))
                            if numFmt_elem is not None:
                                logging.debug(f"Format for ilvl {ilvl} not found, using default format.")
                                return numFmt_elem.get(qn("w:val"))
    logging.warning(f"Numbering format for numId {numId} not found.")
    return None

def get_list_numId(main_doc, desired_fmt, ilvl="0"):
    """
    Creates a new numbering instance in the main document based on an abstract numbering
    definition that uses the desired format (e.g. 'decimal' or 'bullet').
    A level override is added for the specified 'ilvl' so that numbering always starts at 1.
    """
    numbering = main_doc.part.numbering_part.element
    abstractNums = numbering.findall(qn("w:abstractNum"))
    list_abstract_num = None
    for an in abstractNums:
        # Find an abstract definition that has the format we want at the level we want
        for lvl in an.findall(qn("w:lvl")):
            if lvl.get(qn("w:ilvl")) == ilvl:
                numFmt = lvl.find(qn("w:numFmt"))
                if numFmt is not None and numFmt.get(qn("w:val")) == desired_fmt:
                    list_abstract_num = an
                    break
        if list_abstract_num:
            break
            
    # Fallback: If no match at the specific level, find first match for the format
    if list_abstract_num is None:
        for an in abstractNums:
            lvl = an.find(qn("w:lvl"))
            if lvl is not None:
                numFmt = lvl.find(qn("w:numFmt"))
                if numFmt is not None and numFmt.get(qn("w:val")) == desired_fmt:
                    list_abstract_num = an
                    break

    if list_abstract_num is None:
        raise ValueError(
            f"No '{desired_fmt}' numbering definition found in main document"
        )
    abstract_num_id = list_abstract_num.get(qn("w:abstractNumId"))

    # Create a new numbering instance (a new <w:num>) in the main doc.
    nums = numbering.findall(qn("w:num"))
    next_num_id = max([int(num.get(qn("w:numId"))) for num in nums]) + 1 if nums else 1
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(next_num_id))
    abstractNumId_elem = OxmlElement("w:abstractNumId")
    abstractNumId_elem.set(qn("w:val"), abstract_num_id)
    num.append(abstractNumId_elem)
    
    # *** FIX: Apply the level override to the correct 'ilvl' ***
    lvlOverride = OxmlElement("w:lvlOverride")
    lvlOverride.set(qn("w:ilvl"), ilvl) # Use the passed-in ilvl
    startOverride = OxmlElement("w:startOverride")
    startOverride.set(qn("w:val"), "1")
    lvlOverride.append(startOverride)
    num.append(lvlOverride)

    numbering.append(num)
    return next_num_id

def set_numId(element, numId, ilvl="0", additional_indent=360):
    """
    Removes any existing numbering from the paragraph element, sets the numbering id AND level,
    and applies additional indent *only* to the top level (ilvl=0).
    """
    pPr = element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        element.insert(0, pPr)

    # Remove existing numbering
    numPr = pPr.find(qn("w:numPr"))
    if numPr is not None:
        pPr.remove(numPr)
    
    # Add new numbering properties
    numPr = OxmlElement("w:numPr")
    
    ilvl_elem = OxmlElement("w:ilvl")
    ilvl_elem.set(qn("w:val"), ilvl) # <-- **** THE FIX ****
    
    numId_elem = OxmlElement("w:numId")
    numId_elem.set(qn("w:val"), str(numId))
    
    numPr.append(ilvl_elem)
    numPr.append(numId_elem)
    pPr.append(numPr)
    
    # Apply additional indent ONLY to the top level (ilvl="0")
    # Nested levels (ilvl="1", "2", etc.) should get their indentation
    # from the numbering definition itself (which get_list_numId finds).
    if ilvl == "0":
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        
        current_left = ind.get(qn("w:left"))
        # Add our additional indent to the existing left indent
        new_left = (
            additional_indent
            if current_left is None
            else int(current_left) + additional_indent
        )
        ind.set(qn("w:left"), str(new_left))
        
        # Let's also set a default hanging indent for level 0 if not present
        current_hanging = ind.get(qn("w:hanging"))
        if current_hanging is None:
            ind.set(qn("w:hanging"), "360") # Common default hanging indent

def set_paragraph_alignment(element, alignment):
    """
    Sets the paragraph’s alignment to the specified value (e.g., 'left', 'both').
    """
    pPr = element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        element.insert(0, pPr)
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), alignment)

def set_paragraph_spacing(element, after=None, before=None):
    """
    Sets the paragraph's spacing properties in twips (1/20th of a point).
    e.g., after="120" is 6pt space. "0" means no space.
    """
    pPr = element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        element.insert(0, pPr)
    
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
        
    if after is not None:
        spacing.set(qn("w:after"), str(after))
    if before is not None:
        spacing.set(qn("w:before"), str(before))


def set_indent(element, additional_indent=360):
    """
    Adds an additional indent to the paragraph element, preserving existing indentation.
    """
    pPr = element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        element.insert(0, pPr)
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    current_left = ind.get(qn("w:left"))
    new_left = (
        additional_indent
        if current_left is None
        else int(current_left) + additional_indent
    )
    ind.set(qn("w:left"), str(new_left))

def copy_styles_from_reference(main_doc, reference_doc_path):
    """
    Copies ColorCommand style definition from reference document to main document if it doesn't exist.
    This ensures that custom ColorCommand style is preserved when merging.
    """
    if not os.path.exists(reference_doc_path):
        logging.warning(f"Reference document not found at {reference_doc_path}")
        return
    
    reference_doc = Document(reference_doc_path)
    styles_to_copy = ['ColorCommand']
    
    for style_name in styles_to_copy:
        try:
            if style_name in main_doc.styles:
                continue
            reference_style = reference_doc.styles[style_name]
            main_doc.styles.element.append(deepcopy(reference_style.element))
            logging.info(f"Copied {style_name} style from reference doc")
        except KeyError:
            logging.warning(f"{style_name} style not found in reference document")




def replace_placeholder(main_doc_path, output_path, sub_docs, additional_indent=360, preserve_formatting_keys=None):
    """
    Replaces a placeholder in the main document with content from sub-documents.
    Inserts one line space after the section header, preserves sub-document styling,
    adds additional indent to top-level regular paragraphs and bullet points/numbering,
    maintains nested content, and inserts tables as they appear.

    Args:
    preserve_formatting_keys: Set/list of jinja tag keys whose sub-docx content
        should be inserted as-is without any formatting overrides (indent, spacing,
        alignment, numbering remapping). Used for docx_extract sub-documents that
        already carry correct formatting from the source document.
    """
    if preserve_formatting_keys is None:
        preserve_formatting_keys = set()
    else:
        preserve_formatting_keys = set(preserve_formatting_keys)

    main_doc = Document(main_doc_path)

    # Copy styles from reference document once at the beginning
    copy_styles_from_reference(main_doc, reference_docx_path)

    for key, sub_doc_path in sub_docs.items():
        if not os.path.isabs(sub_doc_path):
            sub_doc_path = os.path.abspath(sub_doc_path)
        
        try:
            sub_doc = Document(sub_doc_path)
        except Exception as e:
            logging.error(f"Could not open sub-document: {sub_doc_path}. Error: {e}")
            continue # Skip this sub-doc

        key_pattern = "{{ " + key + "|escape }}"
        skip_formatting = key in preserve_formatting_keys
        for para in reversed(main_doc.paragraphs):
            if key_pattern in para.text:
                logging.info(f"Processing paragraph [{key}] for subdocument generation")
                parent = para._element.getparent()
                index = parent.index(para._element)
                parent.remove(para._element)
                
                empty_para = OxmlElement("w:p")
                parent.insert(index, empty_para)
                
                # *** FIX: The key is now (ilvl, format) ***
                # This groups contiguous list items by their level and style
                local_num_mapping = {}
                
                elements_to_insert = []
                for element in sub_doc.element.body:
                    if element.tag.endswith("p"):  # Paragraph element
                        sub_para = None
                        for p_obj in sub_doc.paragraphs:
                            if p_obj._element == element:
                                sub_para = p_obj
                                break
                        
                        if sub_para is None:
                            logging.warning("Could not find Paragraph object for element.")
                            continue 

                        new_elem = deepcopy(sub_para._element)
                        if skip_formatting:
                            # Preserve original formatting as-is for extract sub-docx
                            elements_to_insert.append(new_elem)
                            continue
                        if is_heading(sub_para):
                            local_num_mapping = {} # Reset list mapping on heading
                            set_paragraph_alignment(new_elem, "left") 
                            set_paragraph_spacing(new_elem, after="120", before="240")
                        elif is_numbered(sub_para):
                            pPr = sub_para._element.pPr
                            sub_orig_numId = None
                            sub_orig_ilvl = "0"
                            sub_num_format = None
                            
                            if pPr is not None:
                                numPr = pPr.find(qn("w:numPr"))
                                if numPr is not None:
                                    ilvl_elem = numPr.find(qn("w:ilvl"))
                                    if ilvl_elem is not None:
                                        sub_orig_ilvl = ilvl_elem.get(qn("w:val"))

                                    sub_numId_elem = numPr.find(qn("w:numId"))
                                    if sub_numId_elem is not None:
                                        sub_orig_numId = sub_numId_elem.get(qn("w:val"))
                                        sub_num_format = get_numbering_format(
                                            sub_doc, sub_orig_numId, sub_orig_ilvl
                                        )
                            
                            if sub_num_format is None:
                                sub_num_format = "decimal" 

                            # *** FIX: Use (ilvl, format) as the key ***
                            num_key = (sub_orig_ilvl, sub_num_format)
                            
                            if num_key not in local_num_mapping:
                                try:
                                    # *** FIX: Pass ilvl to get_list_numId ***
                                    new_list_id = get_list_numId(main_doc, sub_num_format, sub_orig_ilvl)
                                    local_num_mapping[num_key] = new_list_id
                                except ValueError as e:
                                    logging.warning(f"Could not find numbering format '{sub_num_format}' (ilvl {sub_orig_ilvl}) in main doc. Defaulting. Error: {e}")
                                    new_list_id = get_list_numId(main_doc, "decimal", sub_orig_ilvl)
                                    local_num_mapping[num_key] = new_list_id
                            else:
                                new_list_id = local_num_mapping[num_key]
                                
                            set_numId(new_elem, new_list_id, sub_orig_ilvl, additional_indent)
                            set_paragraph_alignment(new_elem, "left") 
                            set_paragraph_spacing(new_elem, after="0")
                        else:
                            # Not a heading, not numbered. Reset mapping for lists.
                            # local_num_mapping = {}
                            set_indent(new_elem, additional_indent)
                            set_paragraph_alignment(new_elem, "left") 
                            set_paragraph_spacing(new_elem, before="0", after="120")
                        
                        elements_to_insert.append(new_elem)

                    elif element.tag.endswith("tbl"):  # Table element
                        # A table also breaks a list
                        local_num_mapping = {}
                        new_table_elem = deepcopy(element)
                        elements_to_insert.append(new_table_elem)

                # Insert all collected elements
                for i, elem_to_insert in enumerate(elements_to_insert):
                    parent.insert(index + i + 1, elem_to_insert) 
                
                break 

    main_doc.save(output_path)
    logging.info("Execution completed for replace_placeholder")
        
def merge_json_arrays(data_array, metadata_array):
    merged_array = []

    for vs_item, meta_item in zip(data_array, metadata_array):
        merged_obj = {}
        for key in vs_item.keys():
            merged_obj[key] = {**vs_item[key], **meta_item.get(key, {})}
        merged_array.append(merged_obj)

    return merged_array


def reform_table_cell(generatedcontent, metadata):
    marged_data = merge_json_arrays(generatedcontent, metadata)
    table_cell_fields = []
    for row in marged_data:
        fieldlabel = ""
        name = ""
        for col_name, col_data in row.items():
            if col_data.get("static", False):
                fieldlabel = col_name
                name = col_data.get("value")
                break
        if not fieldlabel or not name:
            logging.error(f"Invalid row entry {row}")
            continue

        for columnlabel, column_data in row.items():
            # ignore the missing value consider as locked
            if column_data.get("locked", False):
                continue
            # ignore the missing value consider as static, i.e row label ref
            if fieldlabel == columnlabel:
                continue
            column_data.pop("static")
            entry = {
                "name": name,
                "columnlabel": columnlabel,
                "generatedcontent": column_data.pop("value"),
                "fieldlabel": fieldlabel,
            }
            entry.update(column_data)
            table_cell_fields.append(entry)
    return {"fields": table_cell_fields}


def reform_table(generatedcontent, metadata):
    if metadata and metadata.get("need_to_merged_table_column", True) is False:
        entry = dict(**metadata)
        # Sanitize the JSON dictionary keys
        generatedcontent = [sanitize_dict_keys(d) for d in generatedcontent]
        entry["generatedcontent"] = generatedcontent
        return {"fields": [entry]}

    df = pd.DataFrame(generatedcontent)
    data_dict = df.to_dict(orient="list")
    transformed_data = []
    for name, generatedcontent in data_dict.items():
        entry = {"name": name, "generatedcontent": generatedcontent}
        entry.update(metadata.get(name, {}))
        transformed_data.append(entry)

    return {"fields": transformed_data}


def reform_paragraph(generatedcontent, metadata):
    fields = {}
    # check table column is already merged into single entry
    if metadata and isinstance(metadata, dict):
        field = dict(**metadata)
        field["generatedcontent"] = generatedcontent
        transformed_data = {"fields": [field]}
        return transformed_data

    transformed_data = {
        "fields": metadata,
    }
    return transformed_data


def reform_table_and_table_cell_fields(aggregated_config_template):
    reform_data = []
    for section_dict in aggregated_config_template:
        content = section_dict.pop("content")
        modified_content = []
        for item in content:
            if item["type"] == "table_cell":
                generatedcontent = item.pop("generatedcontent")
                metadata = item.pop("metadata")
                new_item = dict(**item)
                new_item.update(reform_table_cell(generatedcontent, metadata))
                modified_content.append(new_item)
            elif item["type"] == "table":
                generatedcontent = item.pop("generatedcontent")
                metadata = item.pop("metadata")
                new_item = dict(**item)
                new_item.update(reform_table(generatedcontent, metadata))
                modified_content.append(new_item)
            elif item["type"] == "paragraph":
                generatedcontent = item.pop("generatedcontent")
                metadata = item.pop("metadata")
                new_item = dict(**item)
                new_item.update(reform_paragraph(generatedcontent, metadata))
                modified_content.append(new_item)
            else:
                modified_content.append(item)

        entry = dict(**section_dict)
        entry["content"] = modified_content
        reform_data.append(entry)
    return reform_data

def remove_tags(text):
    """Remove &amp; tag from the generated html content for rendering"""
    text = text.replace("&amp;", "&")
    return text

def remove_separator_rows(html_content):
    """Remove table rows that contain only --- separators"""
    separator_row_pattern = r'<tr>\s*(?:<td[^>]*>---</td>\s*)+</tr>'
    return re.sub(separator_row_pattern, '', html_content)

def process_table_headers(text):
    """Process tables to bold headers before separators"""
    lines = text.split('\n')
    processed_lines = []
    sep_pattern = r'\|\s*:?---+:?\s*\|'

    for i, line in enumerate(lines):
        if '|' in line and i + 1 < len(lines) and re.search(sep_pattern, lines[i + 1]):
            # This is a header row (line before separator)
            cells = [c.strip() for c in line.split('|')[1:-1]]
            line = '| ' + ' | '.join([f'**{c}**' if c else '' for c in cells]) + ' |'
        processed_lines.append(line)
    
    return '\n'.join(processed_lines)

def fix_text_formatting(text):
    html_tags = r'(?:a|abbr|address|area|article|aside|audio|b|base|bdi|bdo|blockquote|body|br|button|canvas|caption|cite|code|col|colgroup|data|datalist|dd|del|details|dfn|dialog|div|dl|dt|em|embed|fieldset|figcaption|figure|footer|form|h[1-6]|head|header|hr|html|i|iframe|img|input|ins|kbd|label|legend|li|link|main|map|mark|meta|meter|nav|noscript|object|ol|optgroup|option|output|p|param|picture|pre|progress|q|rp|rt|ruby|s|samp|script|section|select|small|source|span|strong|style|sub|summary|sup|svg|table|tbody|td|template|textarea|tfoot|th|thead|time|title|tr|track|u|ul|var|video|wbr)'

    text = re.sub(r'\|\n\n\|', '|\n|', text)                                                   ## In table, re-factors the seprators. |\n\n| > |\n|
    text = re.sub(r'\*\*([^*]+)\*\*', r'\1 ', text)                                            ## Identify and remove double stars if content is wrapped in double stars. **TEXT** > TEXT
    text = process_table_headers(text)                                                         ## Identify and bold table headers
    text = re.sub(rf'<(?!/?{html_tags}(?:/>|>|\s)|{html_tags}/)([^>]+)>', r'&lt;\1&gt;', text) ## Identify and replace < > with &lt; &gt; to avoid markdown conversion. Apart from HTML tags
    text = re.sub(r'^#(.*)$', r'*&#35;&nbsp;\1&nbsp;*', text, flags=re.MULTILINE)              ## Finds lines that starts with # and wrap them with tags(begining and ending)
    text = re.sub(r'```(.*?)```', r'<pre>\1</pre>', text, flags=re.DOTALL)                     ## Finds code blocks wrapped in triple back ticks and converts them to HTML tags <pre>
    text = re.sub(r'\n  ([^\n-].*)', r'<br>&nbsp;&nbsp;\1', text)                              ## Finds line that starts with 2 spaces and replaces the newlinw with <br>
    text = re.sub(r'\*&#35;(&nbsp;| )?([^*]+?)( |&nbsp;)?\*', r'*&#35;&nbsp;\2&nbsp;*', text)  ## Standardizes begining and ending tags
    text = re.sub(r'\*([^*<]+)\*', r'<em>\1</em>', text)                                       ## Convert *...* to <em>...</em> for proper italic rendering inside HTML blocks
    return text

def move_bold_rows_to_thead(html_content):
    """Move rows with bold cells to thead section"""
    bold_row_pattern = r'<tr>.*?<strong>.*?</tr>'
    bold_rows = re.findall(bold_row_pattern, html_content, re.DOTALL)
    
    for row in bold_rows:
        # Convert all td to th which matches bold_row_pattern
        header_row = re.sub(r'<td([^>]*)>(.*?)</td>', r'<th\1>\2</th>', row)
        html_content = html_content.replace(row, '', 1)
        html_content = html_content.replace('</thead>', header_row + '\n</thead>', 1)
    
    return html_content

def merge_empty_th_cells(html_content):
    """Merge consecutive empty table header (th) cells with colspan"""
    pattern = r'(<th[^>]*>.*?</th>)(\s*<th[^>]*>\s*</th>)+'
    
    def replace_func(match):
        full_match = match.group(0)
        first_th = match.group(1)
        total_cells = full_match.count('</th>')
        return first_th.replace('>', f' colspan="{total_cells}">', 1)
    
    return re.sub(pattern, replace_func, html_content)

def process_json_to_subdocx(input_json, output_dir="/data"):
    jinjatag_path_map = {}
    render_counter = 1  # Initialize counter

    for section in input_json:
        for content in section["content"]:
            if content["type"] == "paragraph":
                for field in content["fields"]:
                    if field["generatedContent"] == "":
                        continue
                    jinjatag = field["jinjaTag"]
                    formatted_text = fix_text_formatting(field["generatedContent"])
                    logging.debug(f"After Text formatting :{formatted_text}")
                    html_content = markdown.markdown(formatted_text, extensions=['tables', 'sane_lists'])
                    html_content = move_bold_rows_to_thead(html_content)
                    html_content = merge_empty_th_cells(html_content)
                    html_content = remove_separator_rows(html_content)
                    html_content = remove_tags(html_content)
                    logging.debug(f"HTML Content: {html_content}")
                    docx_path = os.path.join(
                        output_dir, f"tmp_render_{render_counter}.docx"
                    )
                    if os.path.exists(reference_docx_path):
                        logging.info("Converting text using reference document")
                        pypandoc.convert_text(
                        html_content, "docx", format="html", outputfile=docx_path,
                        extra_args=[f"--reference-doc={reference_docx_path}"]
                        )  
                        
                    else:
                        logging.info("Converting text without using reference document")
                        pypandoc.convert_text(
                            html_content, "docx", format="html", outputfile=docx_path
                        )
                    # Store mapping and increment counter
                    jinjatag_path_map[jinjatag] = docx_path
                    render_counter += 1
    return jinjatag_path_map
