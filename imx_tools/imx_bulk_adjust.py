import sys
from datetime import datetime
from pathlib import Path
import re

import pandas as pd
from lxml import etree
from lxml.etree import Element
import xmlschema
from loguru import logger


logger.info("reading xsd")
XSD_IMX_V124 = xmlschema.XMLSchema('xsd_1.2.4/IMSpoor-1.2.4-Communication.xsd')
logger.success("xsd reading finished")


def log_filter(record):
    return record["level"].name in ["SUCCESS", "ERROR"]


logger.remove(0)
logger.add(sys.stdout, level="DEBUG", filter=log_filter)


def get_all_elements_by_name(element: Element, element_name: str) -> list[Element]:
    return element.xpath(f'.//*[local-name()="{element_name}"]')


def get_elements_by_name(element: Element, element_name: str) -> list[Element]:
    return element.xpath(f'./*[local-name()="{element_name}"]')


def get_all_elements_containing_attribute(element: Element, attribute_name: str, value: str | None = None) -> list[Element]:
    if value:
        return element.xpath(f'..//*[{attribute_name}="{value}"]')
    return element.xpath(f'..//*[{attribute_name}]')


def set_attribute(element: Element, attribute_name: str, value: str):
    old_value = element.get(attribute_name.replace("@", ""))
    element.set(attribute_name.replace("@", ""), value)
    parent = element.getparent()

    if parent is not None:
        tester = parent.index(element)
        parent.insert(tester, etree.Comment(f"Attribute element below changed {attribute_name} old: {old_value} new: {value}"))


def set_element_text(element: Element, element_name: str, value: str):
    temp = get_elements_by_name(element, element_name)
    if temp:
        old_value = temp[0].text
        temp[0].text = value
        parent = temp[0].getparent()
        if parent is not None:
            test = parent.index(temp[0])
            parent.insert(test, etree.Comment(f"Element text below changed {element_name} old: {old_value} new: {value}"))


def get_parent_and_target(element: Element, path_split: list[str]) -> tuple[Element, str]:
    parent = element
    for idx, element_name in enumerate(path_split[:-1]):
        elements = get_elements_by_name(parent, element_name)
        if len(elements) > 1:
            try:
                parent = elements[int(path_split[idx+1])]
            except Exception as e:
                raise ValueError(f'{".".join(path_split)} index "{int(path_split[idx+1])}" out of range')
        else:
            if elements:
                parent = elements[0]
    return parent, path_split[-1]


def handle_attribute(parent: Element, attribute_name: str, value: str, old_value: str | None):
    attr_key = attribute_name.replace("@", "")
    if old_value:
        if parent.attrib.get(attr_key) == old_value:
            set_attribute(parent, attribute_name, value)
        elif not parent.attrib.get(attr_key):
            raise ValueError(f"Attribute not found: {attr_key}")
        else:
            raise ValueError(f"Attribute mismatch: {attr_key} has value {parent.attrib.get(attr_key)}")
    else:
        if attr_key not in parent.attrib:
            set_attribute(parent, attribute_name, value)


def handle_element(parent: Element, element_name: str, value: str, old_value: str | None):
    elements = get_elements_by_name(parent, element_name)
    if elements and elements[0].text == old_value:
        set_element_text(parent, element_name, value)
    else:
        raise ValueError("Mismatch in old value for element text.")


def set_attribute_or_element_by_path(puic_object: Element, path: str, value: str, old_value: str | None):
    path_split = path.split(".")
    if path_split[-1].startswith("@"):  # Attribute case
        parent, attribute_name = get_parent_and_target(puic_object, path_split)
        handle_attribute(parent, attribute_name, value, old_value)
    else:  # Element case
        parent, element_name = get_parent_and_target(puic_object, path_split)
        handle_element(parent, element_name, value, old_value)


def delete_attribute_if_matching(puic_object: Element, path: str, value: str):
    path_split = path.split(".")
    attribute_name = path_split[-1]
    if not attribute_name.startswith("@"):  # Ensure it's an attribute
        raise ValueError("Path must end with an attribute (e.g., '@id').")

    parent, _ = get_parent_and_target(puic_object, path_split[:-1])
    attribute_name = attribute_name.replace("@", "")
    if parent.attrib.get(attribute_name) == value:
        del parent.attrib[attribute_name]
    else:
        raise ValueError(f"Attribute '{attribute_name}' value does not match '{value}'.")


def delete_element(element: Element):
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)

def set_metadata(node:Element):
    timestamp = datetime(2025, 6, 29, 23, 59, 59).strftime('%Y-%m-%dT%H:%M:%SZ')
    # timestamp = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    metadata = node.find('.//{http://www.prorail.nl/IMSpoor}Metadata')

    source_value = metadata.get("source")
    source_value_to_set = "ProRail_DV"
    if source_value and not source_value.endswith(source_value_to_set):
        if source_value.endswith("ProRail"):
            source_value = source_value + f"_DV"
        else:
            source_value = source_value + f"_{source_value_to_set}"

    metadata.set("source", source_value)
    metadata.set("lifeCycleStatus", "FinalDesign")
    metadata.set("originType", "Unknown")

    metadata.set("registrationTime", timestamp)
    _ = metadata.getparent()
    test = _.index(metadata)
    _.insert(test, etree.Comment(f"Metadata changed"))



def set_source_attribute(node: Element):
    puic_ = node.get('puic')
    if puic_:
        set_metadata(node)
        logger.success(f"metadata for {puic_} set")

    parent = node.getparent()
    puic_ = parent.get('puic')
    while parent is not None:
        if parent.tag == '{http://www.prorail.nl/IMSpoor}Project' or parent.tag == '{http://www.prorail.nl/IMSpoor}Situation':
            break
        elif puic_ is not None:
            set_metadata(parent)

            logger.success(f"metadata for {puic_} set")
        parent = parent.getparent()


def process_changes(change_items: list[dict], puic_dict: dict[str, Element]):
    for change in change_items:
        logger.info(f"processing change {change}")

        puic = change["puic"]
        if puic not in puic_dict:
            change["status"] = f"object not present: {puic}"
            continue

        imx_object_element: Element = puic_dict[puic]

        object_type = change["ObjectType"]
        operation = change["Operation"]

        try:
            if imx_object_element.tag != f"{{http://www.prorail.nl/IMSpoor}}{object_type.split('.')[-1]}":
                raise ValueError(
                    f"Object tag {object_type} does not match tag of found object {imx_object_element.tag.split('}')[1]}"
                )

            match operation:
                case "CreateAttribute":
                    set_attribute_or_element_by_path(imx_object_element, change["Atribute"], f"{change['Waarde nieuw']}", None)
                    set_source_attribute(imx_object_element)
                    change["status"] = "processed"

                case "UpdateAttribute":
                    set_attribute_or_element_by_path(imx_object_element, change["Atribute"], f"{change['Waarde nieuw']}", f"{change['Waarde oud']}")
                    set_source_attribute(imx_object_element)
                    change["status"] = "processed"

                case "DeleteAttribute":
                    delete_attribute_if_matching(imx_object_element, change["Atribute"], change["Waarde oud"])
                    change["status"] = "processed"

                case "DeleteObject":
                    delete_element(imx_object_element)
                    change["status"] = "processed"
                    puic_dict.pop(puic)  # Remove deleted object from the dictionary

                case _:
                    change["status"] = f"NOT processed: {operation} is not a valid operation"

        except Exception as e:
            logger.error(e)
            change["status"] = f"Error: {e}"

        finally:
            errors = list(XSD_IMX_V124.iter_errors(imx_object_element))
            if errors:
                change["status"] = f"{change['status']} - XSD invalid!"
                change["xsd_errors"] = "".join([error.reason for error in errors])
                logger.error(change["xsd_errors"])

        logger.success(f"processing change {change} done")


if __name__ == "__main__":
    # input
    xml_file = r"C:\test data bulk verbetering\zeeuws_vlaanderen.xml"
    excel_file = r"C:\test data bulk verbetering\Updateslijst_Zeeuws_Vlaanderen.xlsx"
    excel_sheet = "Blad1"

    # output
    output_xml_file = r"C:\test data bulk verbetering\output_file_path.xml"
    output_excel_file = r"C:\test data bulk verbetering\output_file_path.xlsx"

    logger.info("loading xml")
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(xml_file, parser=parser)
    logger.success("loading xml finished")

    # todo: refactor to use puic dict for speed
    puic_objects = tree.findall(".//*[@puic]")
    puic_dict = {value.get("puic"): value for value in puic_objects}

    df = pd.read_excel(excel_file, sheet_name=excel_sheet)
    df = df.fillna("None")

    change_items = df.to_dict(orient="records")

    logger.info("processing xml")
    process_changes(change_items, puic_dict)
    logger.success("processing xml finshed")

    tree.write(output_xml_file, encoding="UTF-8", pretty_print=True)
    pd.DataFrame(change_items).to_excel(output_excel_file, index=False)
