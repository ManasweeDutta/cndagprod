from copy import deepcopy
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
import os
import logging


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
                # IMPORTANT:
                # Insert EXACTLY at placeholder location
                # to preserve original indentation
                #
                index = parent.index(
                    para._element
                )

                #
                # Insert RAW OOXML blocks
                #
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

                        parent.insert(
                            index,
                            new_elem
                        )

                        index += 1

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

                        parent.insert(
                            index,
                            new_tbl
                        )

                        index += 1

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
