from openpyxl import Workbook, load_workbook
import xml.etree.ElementTree as ET
import re

COUNTRIES_OUTPUT_FILE = "countries_output.xml"
GOODS_OUTPUT_FILE = "goods_output.xml"

wb = load_workbook('master_data.xlsx')

countries = ET.Element("Countries")
countries_tree = ET.ElementTree(countries)

goods = ET.Element("Goods")
goods_tree = ET.ElementTree(goods)


def country_exists(country_name):
    for country in countries.findall("Country"):
        name = country.find("Name").text
        if name == country_name:
            return country
    return None


def country_profiles(country, row):
    advancement_lvl = row[2]
    description = row[4]

    ET.SubElement(country, "Advancement_Level").text = advancement_lvl
    ET.SubElement(country, "Description").text = description


def statistics_on_children(country, row):
    stats = country.find("Country_Statistics")
    if stats == None:
        stats = ET.SubElement(country, "Country_Statistics")

    regex = r"(^\d+(,\d+)*(\.\d+(e\d+)?)?)(\s\((\d+(,\d+)*(\.\d+(e\d+)?)?)\))?$"

    stat_type = row[2]
    sector = row[3]
    age = row[4]
    percent = row[5]
    match = re.match(regex, str(percent))

    if stat_type == "Working (% and population)" or stat_type == "Working children by sector":
        child_work = stats.find("Children_Work_Statistics")
        if child_work == None:
            child_work = ET.SubElement(stats, "Children_Work_Statistics")

        if stat_type == "Working (% and population)":
            ET.SubElement(child_work, "Age_Range").text = age
            ET.SubElement(
                child_work, "Total_Percentage_of_Working_Children").text = match.group(1) if match else ""
            ET.SubElement(
                child_work, "Total_Working_Population").text = match.group(6) if match else ""
        elif stat_type == "Working children by sector":
            ET.SubElement(child_work, sector).text = match.group(
                1) if match else ""
    elif stat_type == "Attending School (%)":
        education = stats.find("Education_Statistics_Attendance_Statistics")
        if education == None:
            education = ET.SubElement(
                stats, "Education_Statistics_Attendance_Statistics")

        ET.SubElement(education, "Age_Range").text = age

        if match:
            ET.SubElement(
                education, "Percentage").text = match.group(1)
    elif stat_type == "Combining Work and School (%)":
        work_and_school = stats.find(
            "Children_Working_and_Studying_7-14_yrs_old")
        if work_and_school == None:
            work_and_school = ET.SubElement(
                stats, "Children_Working_and_Studying_7-14_yrs_old")

        ET.SubElement(work_and_school, "Age_Range").text = age

        if match:
            ET.SubElement(
                work_and_school, "Total").text = match.group(1)
    elif stat_type == "Primary Completion Rate (%)":
        completion_rate = stats.find(
            "UNESCO_Primary_Completion_Rate")
        if completion_rate == None:
            completion_rate = ET.SubElement(
                stats, "UNESCO_Primary_Completion_Rate")
            ET.SubElement(
                completion_rate, "Rate").text = match.group(1) if match else ""


def read_row(country, row, ws_idx):
    if ws_idx == 1:
        country_profiles(country, row)
    elif ws_idx == 2:
        statistics_on_children(country, row)


for idx, sheet in enumerate(wb.sheetnames):
    if sheet == "Instructions":
        continue
    ws = wb[sheet]
    for row in ws.iter_rows(min_row=2, values_only=True):
        country_name = row[1]
        country = country_exists(country_name)
        if country == None:
            country = ET.SubElement(countries, "Country")
            name = ET.SubElement(country, "Name")
            name.text = country_name

        read_row(country, row, idx)


countries_tree.write(COUNTRIES_OUTPUT_FILE)
