from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
import os
import logging


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def iter_block_items(parent):
    """
    Yield paragraphs and tables in document order
    """

    parent_elm = parent.element.body

    for child in parent_elm.iterchildren():

        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)

        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def copy_styles(source_doc, target_doc):
    """
    Copy ONLY missing styles from source document
    into target document.

    Do NOT overwrite existing styles.
    """

    target_styles = target_doc.styles

    for style in source_doc.styles:

        try:

            #
            # Skip existing styles
            #
            target_styles[style.name]

        except KeyError:

            try:

                target_styles.element.append(
                    deepcopy(style._element)
                )

            except Exception as e:

                logging.warning(
                    f"Failed to copy style {style.name}: {e}"
                )


def copy_numbering_part(source_doc, target_doc):
    """
    Copy numbering definitions from source doc
    into target doc using NEW unique IDs
    to avoid numbering collisions.
    """

    try:

        src_numbering = (
            source_doc.part.numbering_part._element
        )

        tgt_numbering = (
            target_doc.part.numbering_part._element
        )

        #
        # Existing IDs in target
        #
        existing_num_ids = set()
        existing_abs_ids = set()

        for child in tgt_numbering:

            tag = child.tag.split("}")[-1]

            if tag == "num":

                num_id = child.get(
                    f"{{{W_NS}}}numId"
                )

                if num_id:
                    existing_num_ids.add(
                        int(num_id)
                    )

            elif tag == "abstractNum":

                abs_id = child.get(
                    f"{{{W_NS}}}abstractNumId"
                )

                if abs_id:
                    existing_abs_ids.add(
                        int(abs_id)
                    )

        #
        # Generate safe IDs
        #
        next_num_id = (
            max(existing_num_ids, default=0) + 100
        )

        next_abs_id = (
            max(existing_abs_ids, default=0) + 100
        )

        abs_mapping = {}
        num_mapping = {}

        #
        # Copy abstractNum definitions
        #
        for child in src_numbering:

            tag = child.tag.split("}")[-1]

            if tag != "abstractNum":
                continue

            old_abs_id = int(
                child.get(
                    f"{{{W_NS}}}abstractNumId"
                )
            )

            new_abs_id = next_abs_id
            next_abs_id += 1

            abs_mapping[
                old_abs_id
            ] = new_abs_id

            new_child = deepcopy(child)

            new_child.set(
                f"{{{W_NS}}}abstractNumId",
                str(new_abs_id)
            )

            tgt_numbering.append(
                new_child
            )

        #
        # Copy num definitions
        #
        for child in src_numbering:

            tag = child.tag.split("}")[-1]

            if tag != "num":
                continue

            old_num_id = int(
                child.get(
                    f"{{{W_NS}}}numId"
                )
            )

            new_num_id = next_num_id
            next_num_id += 1

            num_mapping[
                old_num_id
            ] = new_num_id

            new_child = deepcopy(child)

            new_child.set(
                f"{{{W_NS}}}numId",
                str(new_num_id)
            )

            #
            # Remap abstractNumId
            #
            abs_ref = new_child.find(
                ".//w:abstractNumId",
                {
                    "w": W_NS
                }
            )

            if abs_ref is not None:

                old_abs = int(
                    abs_ref.get(
                        f"{{{W_NS}}}val"
                    )
                )

                abs_ref.set(
                    f"{{{W_NS}}}val",
                    str(abs_mapping[old_abs])
                )

            tgt_numbering.append(
                new_child
            )

        return num_mapping

    except Exception as e:

        logging.warning(
            f"Failed to copy numbering part: {e}"
        )

        return {}


def clean_heading_numbering(paragraph_element):
    """
    Remove numbering ONLY from heading paragraphs
    so the master template controls heading numbering.
    """

    pStyle = paragraph_element.find(
        ".//w:pStyle",
        {
            "w": W_NS
        }
    )

    if pStyle is None:
        return

    style_val = (
        pStyle.get(
            f"{{{W_NS}}}val",
            ""
        ).lower()
    )

    #
    # Detect heading styles
    #
    if "heading" not in style_val:
        return

    #
    # Remove numbering from heading ONLY
    #
    pPr = paragraph_element.find(
        ".//w:pPr",
        {
            "w": W_NS
        }
    )

    if pPr is None:
        return

    numPr = pPr.find(
        ".//w:numPr",
        {
            "w": W_NS
        }
    )

    if numPr is not None:

        pPr.remove(numPr)


def remap_paragraph_numbering(
    paragraph_element,
    num_mapping
):
    """
    Remap paragraph numbering IDs
    to prevent numbering collisions.
    """

    pPr = paragraph_element.find(
        ".//w:numPr",
        {
            "w": W_NS
        }
    )

    if pPr is None:
        return

    numId = pPr.find(
        ".//w:numId",
        {
            "w": W_NS
        }
    )

    if numId is None:
        return

    old_id = int(
        numId.get(
            f"{{{W_NS}}}val"
        )
    )

    if old_id in num_mapping:

        numId.set(
            f"{{{W_NS}}}val",
            str(num_mapping[old_id])
        )


def replace_placeholder(
    main_doc_path,
    output_path,
    sub_docs,
    additional_indent=360,
    preserve_formatting_keys=None
):
    """
    TRUE RAW OOXML MERGE

    Preserves:
    - exact indentation
    - exact bullets
    - exact numbering
    - exact heading numbering
    - exact spacing
    - exact bash formatting
    - exact tables
    - exact paragraph styling
    """

    if preserve_formatting_keys is None:
        preserve_formatting_keys = []

    logging.info(
        f"Loading main document: {main_doc_path}"
    )

    #
    # Load main document
    #
    main_doc = Document(
        main_doc_path
    )

    #
    # Process placeholders
    #
    for placeholder_key, sub_doc_path in sub_docs.items():

        logging.info(
            f"Processing placeholder: {placeholder_key}"
        )

        #
        # Validate subdoc path
        #
        if not os.path.exists(sub_doc_path):

            logging.warning(
                f"Subdoc not found: {sub_doc_path}"
            )

            continue

        #
        # Placeholder text
        #
        placeholder = (
            f"{{{{ {placeholder_key}|escape }}}}"
        )

        #
        # Load sub document
        #
        sub_doc = Document(
            sub_doc_path
        )

        #
        # Copy styles ONLY for non-preserved docs
        #
        if (
            placeholder_key
            not in preserve_formatting_keys
        ):

            copy_styles(
                sub_doc,
                main_doc
            )

        #
        # Copy numbering definitions safely
        #
        num_mapping = copy_numbering_part(
            sub_doc,
            main_doc
        )

        placeholder_found = False

        #
        # Find placeholder paragraph
        #
        for para in main_doc.paragraphs:

            if placeholder in para.text:

                placeholder_found = True

                logging.info(
                    f"Found placeholder: {placeholder}"
                )

                parent = para._element.getparent()

                #
                # Insert EXACTLY at placeholder location
                #
                index = parent.index(
                    para._element
                )

                #
                # Collect all elements first
                #
                elements_to_insert = []

                for block in iter_block_items(sub_doc):

                    #
                    # Paragraph
                    #
                    if isinstance(
                        block,
                        Paragraph
                    ):

                        new_elem = deepcopy(
                            block._element
                        )

                        #
                        # Remove heading numbering
                        # so template numbering applies
                        #
                        clean_heading_numbering(
                            new_elem
                        )

                        #
                        # Remap numbering IDs
                        # for body lists/bullets
                        #
                        remap_paragraph_numbering(
                            new_elem,
                            num_mapping
                        )

                        elements_to_insert.append(
                            new_elem
                        )

                    #
                    # Table
                    #
                    elif isinstance(
                        block,
                        Table
                    ):

                        new_tbl = deepcopy(
                            block._element
                        )

                        elements_to_insert.append(
                            new_tbl
                        )

                #
                # Insert sequentially
                #
                for offset, elem in enumerate(
                    elements_to_insert
                ):

                    parent.insert(
                        index + offset,
                        elem
                    )

                #
                # Remove placeholder LAST
                #
                parent.remove(
                    para._element
                )

                break

        #
        # Placeholder not found
        #
        if not placeholder_found:

            logging.warning(
                f"Placeholder not found: {placeholder}"
            )

    #
    # Save merged document
    #
    main_doc.save(
        output_path
    )

    logging.info(
        f"Merged document saved: {output_path}"
    )
