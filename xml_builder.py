from openpyxl import Workbook, load_workbook
import xml.etree.ElementTree as ET
import re

COUNTRIES_OUTPUT_FILE = "countries_output.xml"
GOODS_OUTPUT_FILE = "goods_output.xml"
year = "2018"

wb = load_workbook('master_data.xlsx')

countries = ET.Element("Countries")
countries_tree = ET.ElementTree(countries)

goods = ET.Element("Goods")
goods_tree = ET.ElementTree(goods)

country_display_names = {
        "RS": "Republika Srpska",
        "FBiH": "Federation of Bosnia and Herzegovina",
        "BD": "Brčko District",
        "BiH": "Bosnia and Herzegovina"
    }


def country_exists(country_name):
    for country in countries.findall("Country"):
        name = country.find("Name").text
        if name == country_name:
            return country
    return None


def check_multiple_territories(country, related_entity):
    territories = country.find("Multiple_Territories")
    if territories == None:
        territories = ET.SubElement(country, "Multiple_Territories")

    if related_entity:
        territories.text = "Yes"
    else:
        territories.text = "No"


def country_profiles(country, row):
    region = row[2]
    advancement_lvl = row[4]
    description = row[6]

    ET.SubElement(country, "Region").text = region
    ET.SubElement(country, "Multiple_Territories")
    ET.SubElement(country, "Advancement_Level").text = advancement_lvl
    ET.SubElement(country, "Description").text = description.strip()

    # create these tags now to be used in later sheets
    ET.SubElement(country, "Goods")


def goods_list(country, row):
    good = row[2]
    child_labor = "Yes" if row[3] == 1 else "No"
    forced_labor = "Yes" if row[4] == 1 else "No"
    forced_child_labor = "Yes" if row[5] == 1 else "No"

    sectors = {"manu": "Manufacturing",
               "mine": "Mining",
               "agri": "Agriculture",
               "other": "Other"}
    sector = sectors[row[6]] if row[6] in sectors else ""
    if sector:
        # countries.xml
        goodsTag = country.find("Goods")
        if goodsTag == None:
            goodsTag = ET.SubElement(country, "Goods")

        goodTag = ET.SubElement(goodsTag, "Good")
        ET.SubElement(goodTag, "Good_Name").text = good
        ET.SubElement(goodTag, "Child_Labor").text = child_labor
        ET.SubElement(goodTag, "Forced_Labor").text = forced_labor
        ET.SubElement(goodTag, "Forced_Child_Labor").text = forced_child_labor

        # goods.xml
        goodTag = None
        for val in goods.findall("Good"):
            name = val.find("Good_Name")
            if name.text == good:
                goodTag = val
                break

        if goodTag == None:
            goodTag = ET.SubElement(goods, "Good")
            ET.SubElement(goodTag, "Good_Name").text = good
            ET.SubElement(goodTag, "Good_Sector").text = sector

        countriesTag = goodTag.find("Countries")
        if countriesTag == None:
            countriesTag = ET.SubElement(goodTag, "Countries")
        countryTag = ET.SubElement(countriesTag, "Country")
        countryName = country.find("Name")
        countryRegion = country.find("Region")
        ET.SubElement(
            countryTag, "Country_Name").text = countryName.text if not countryName == None else ""
        ET.SubElement(
            countryTag, "Country_Region").text = countryRegion.text if not countryRegion == None else ""
        ET.SubElement(countryTag, "Child_Labor").text = child_labor
        ET.SubElement(countryTag, "Forced_Labor").text = forced_labor
        ET.SubElement(
            countryTag, "Forced_Child_Labor").text = forced_child_labor


def statistics_on_children(country, row):
    stats = country.find("Country_Statistics")
    if stats == None:
        stats = ET.SubElement(country, "Country_Statistics")

    regex = r"(^\d+(,\d+)*(\.\d+(e\d+)?)?)(\s\((\d+(,\d+)*(\.\d+(e\d+)?)?)\))?$"

    related_entity = row[2]
    stat_type = row[3]
    sector = row[4]
    age = row[5]
    percent = row[6]
    match = re.match(regex, str(percent))

    age_range = age.replace("to", "-").replace(" ", "") if age else ""
    group = match.group(1) if match else ""
    percentage = str(round(float(group) / 100, 3)
                     ) if is_number(group) else "Unavailable"

    check_multiple_territories(country, related_entity)

    if stat_type == "Working (% and population)" or stat_type == "Working children by sector":
        child_work = stats.find("Children_Work_Statistics")
        if child_work == None:
            child_work = ET.SubElement(stats, "Children_Work_Statistics")

        if stat_type == "Working (% and population)":
            total_work_pop = match.group(6) if match else ""
            if total_work_pop:
                total_work_pop = total_work_pop.replace(",", "")
            
            ET.SubElement(child_work, "Age_Range").text = age_range
            ET.SubElement(
                child_work, "Total_Percentage_of_Working_Children").text = percentage
            ET.SubElement(
                child_work, "Total_Working_Population").text = total_work_pop
        elif stat_type == "Working children by sector" and sector:
            ET.SubElement(child_work, sector).text = percentage
    elif stat_type == "Attending School (%)":
        education = stats.find("Education_Statistics_Attendance_Statistics")
        if education == None:
            education = ET.SubElement(
                stats, "Education_Statistics_Attendance_Statistics")

        ET.SubElement(education, "Age_Range").text = age_range
        ET.SubElement(
            education, "Percentage").text = percentage
    elif stat_type == "Combining Work and School (%)":
        work_and_school = stats.find(
            "Children_Working_and_Studying_7-14_yrs_old")
        if work_and_school == None:
            work_and_school = ET.SubElement(
                stats, "Children_Working_and_Studying_7-14_yrs_old")

        ET.SubElement(work_and_school, "Age_Range").text = age_range
        ET.SubElement(
            work_and_school, "Total").text = percentage
    elif stat_type == "Primary Completion Rate (%)":
        completion_rate = stats.find(
            "UNESCO_Primary_Completion_Rate")
        if completion_rate == None:
            completion_rate = ET.SubElement(
                stats, "UNESCO_Primary_Completion_Rate")
            ET.SubElement(
                completion_rate, "Rate").text = percentage


def ratification_of_international(country, row):
    conventions = country.find("Conventions")
    if conventions == None:
        conventions = ET.SubElement(country, "Conventions")

    convention = row[3]
    ratification = row[4]

    tags = {"ILO C. 138, Minimum Age": "C_138_Ratified", "UN CRC": "Convention_on_the_Rights_of_the_Child_Ratified",
            "ILO C. 182, Worst Forms of Child Labor": "C_182_Ratified",
            "UN CRC Optional Protocol on the Sale of Children, Child Prostitution and Child Pornography": "CRC_Commercial_Sexual_Exploitation_of_Children_Ratified",
            "UN CRC Optional Protocol on Armed Conflict": "CRC_Armed_Conflict_Ratified",
            "Palermo Protocol on Trafficking in Persons": "Palermo_Ratified"}

    if ratification == "1":
        ratification = "Yes"
    elif ratification == "0":
        ratification = "No"
    if convention:
        ET.SubElement(conventions, tags[convention]).text = ratification


def laws_and_regulations(country, row):
    legal = country.find("Legal_Standards")
    if legal == None:
        legal = ET.SubElement(country, "Legal_Standards")

    related_entity = row[2]
    standard = row[3]
    meets_intl_stds = row[5]
    age = row[7]
    calced_age = "Yes" if row[8] == "TRUE" else "No"
    multiple_territories = True if country.find("Multiple_Territories").text == "Yes" else False

    tags = {
        "Compulsory Education Age": "Compulsory_Education",
        "Free Public Education": "Free_Public_Education",
        "Identification of Hazardous Occupations or Activities Prohibited for Children": "Types_Hazardous_Work",
        "Minimum Age for Hazardous Work": "Minimum_Hazardous_Work",
        "Minimum Age for Voluntary State Military Recruitment": "Minumum_Voluntary_Military",
        "Minimum Age for Work": "Minimum_Work",
        "Prohibition of Child Trafficking": "Prohibition_Child_Trafficking",
        "Prohibition of Commercial Sexual Exploitation of Children": "Prohibition_CSEC",
        "Prohibition of Compulsory Recruitment of Children by (State) Military": "Minimum_Compulsory_Military",
        "Prohibition of Forced Labor": "Prohibition_Forced_Labor",
      #  "Prohibition of Military Recruitment": "",
        "Prohibition of Military Recruitment by Non-state Armed Groups": "Minumum_Non_State_Military",
        "Prohibition of Using Children in Illicit Activities": "Prohibition_Illicit_Activities"
    }

    if standard and standard in tags:
        tag = legal.find(tags[standard])
        if tag == None:
            tag = ET.SubElement(legal, tags[standard])
        if multiple_territories:
            tag = ET.SubElement(tag, "Territory")
            display_name = "All Territories"
            if related_entity:
                display_name = country_display_names[related_entity] if related_entity in country_display_names else related_entity
            ET.SubElement(tag, "Territory_Name").text = display_name
            ET.SubElement(tag, "Territory_Display_Name").text = display_name
        ET.SubElement(tag, "Standard").text = meets_intl_stds
        ET.SubElement(tag, "Age").text = age
        ET.SubElement(tag, "Calculated_Age").text = calced_age
        ET.SubElement(tag, "Conforms_To_Intl_Standard").text = meets_intl_stds


def labor_law_enforcement(country, row):
    enforcements = country.find("Enforcements")
    if enforcements == None:
        enforcements = ET.SubElement(country, "Enforcements")

    related_entity = row[2]
    overview = row[3]
    sub_overview = row[4]
    current_year_data = row[5]
    multiple_territories = True if country.find("Multiple_Territories").text == "Yes" else False

    tags = {
        "Labor Inspectorate Funding": "Labor_Funding",
        "Number of Labor Inspectors": "Labor_Inspectors",
        "Inspectorate Authorized to Assess Penalties": "Authorized_Access_Penalties",
        "Initial Training for New Labor Inspectors":
            {"NA": "Labor_New_Employee_Training",
             "Training on New Laws Related to Child Labor": "Labor_New_Law_Training",
             "Refresher Courses Provided": "Labor_Refresher_Courses"},
        "Number of Labor Inspections Conducted":
            {"NA": "Labor_Inspections",
             "Number Conducted at Worksite": "Labor_Worksite_Inspections"},
        "Number of Child Labor Violations Found":
            {"NA": "Labor_Violations",
             "Number of Child Labor Violations for Which Penalties Were Imposed": "Labor_Penalties_Imposed",
             "Number of Child Labor Penalties Imposed that Were Collected": "Labor_Penalties_Collected"},
        "Routine Inspections Conducted":
            {"NA": "Labor_Routine_Inspections_Conducted",
             "Routine Inspections Targeted": "Labor_Routine_Inspections_Targeted"},
        "Unannounced Inspections Permitted":
            {"NA": "Labor_Unannounced_Inspections_Premitted",
             "Unannounced Inspections Conducted": "Labor_Unannounced_Inspections_Conducted"},
        "Complaint Mechanism Exists": "Labor_Complaint_Mechanism",
        "Reciprocal Referral Mechanism Exists Between Labor Authorities and Social Services": "Labor_Referral_Mechanism"
    }

    if overview:
        tag = None
        if isinstance(tags[overview], dict):
            if not sub_overview:
                sub_overview = "NA"
            tag = enforcements.find(tags[overview][sub_overview])
            if tag == None:
                tag = ET.SubElement(
                    enforcements, tags[overview][sub_overview])
        else:
            tag = enforcements.find(tags[overview])
            if tag == None:
                tag = ET.SubElement(
                    enforcements, tags[overview])
        if multiple_territories:
            tag = ET.SubElement(tag, "Territory")
            display_name = "All Territories"
            if related_entity:
                display_name = country_display_names[related_entity] if related_entity in country_display_names else related_entity
            ET.SubElement(tag, "Territory_Name").text = display_name
            ET.SubElement(tag, "Territory_Display_Name").text = display_name
            ET.SubElement(tag, "Enforcement").text = current_year_data
        else:
            tag.text = current_year_data


def criminal_law_enforcement(country, row):
    enforcements = country.find("Enforcements")
    if enforcements == None:
        enforcements = ET.SubElement(country, "Enforcements")

    related_entity = row[2]
    overview = row[3]
    sub_overview = row[4]
    current_year_data = row[5]
    multiple_territories = True if country.find("Multiple_Territories").text == "Yes" else False

    tags = {
        "Number of Investigations": "Criminal_Investigations",
        "Number of Violations Found": "Criminal_Violations",
        "Number of Prosecutions Initiated": "Criminal_Prosecutions",
        "Number of Convictions": "Criminal_Convictions",
        "Initial Training for New Criminal Investigators":
            {"NA": "Criminal_New_Employee_Training",
             "Training on New Laws Related to the Worst Forms of Child Labor": "Criminal_New_Law_Training",
             "Refresher Courses Provided": "Criminal_Refresher_Courses"},
        "Reciprocal Referral Mechanism Exists Between Criminal Authorities and Social Services": "Criminal_Referral_Mechanism"
    }

    if overview and overview in tags:
        tag = None
        if isinstance(tags[overview], dict):
            if not sub_overview:
                sub_overview = "NA"
            tag = enforcements.find(tags[overview][sub_overview])
            if tag == None:
                tag = ET.SubElement(
                    enforcements, tags[overview][sub_overview])
        else:
            tag = enforcements.find(tags[overview])
            if tag == None:
                tag = ET.SubElement(
                    enforcements, tags[overview])
        if multiple_territories:
            tag = ET.SubElement(tag, "Territory")
            display_name = "All Territories"
            if related_entity:
                display_name = country_display_names[related_entity] if related_entity in country_display_names else related_entity
            ET.SubElement(tag, "Territory_Name").text = display_name
            ET.SubElement(tag, "Territory_Display_Name").text = display_name
            ET.SubElement(tag, "Enforcement").text = current_year_data
        else:
            tag.text = current_year_data


def government_actions(country, row):
    actions = country.find("Suggested_Actions")
    if actions == None:
        actions = ET.SubElement(country, "Suggested_Actions")

    area = row[3]
    action = row[4]

    tags = {
        "Legal Framework": "Legal_Framework",
        "Enforcement": "Enforcement",
        "Coordination": "Coordination",
        "Government Policies": "Government_Policies",
        "Social Programs": "Social_Programs"
    }

    if area and action:
        tag = actions.find(tags[area])
        if tag == None:
            tag = ET.SubElement(actions, tags[area])

        action_tag = ET.SubElement(tag, "Action")
        ET.SubElement(action_tag, "Name").text = action.strip()


def deliberative_data(country, row):
    mechanisms = country.find("Mechanisms")
    if mechanisms == None:
        mechanisms = ET.SubElement(country, "Mechanisms")

    yes_no_na = row[2]
    program = row[3]

    tags = {
        "Does the government have a program to combat child labor?": "Program",
        "Does the government have a policy to combat child labor?": "Policy",
        "Does the government have a mechanism to coordinate efforts against CL?": "Coordination"
    }

    if program:
        ET.SubElement(mechanisms, tags[program]).text = yes_no_na


def read_row(country, row, ws_idx):
    options = {1: country_profiles,
               2: statistics_on_children,
               4: ratification_of_international,
               5: laws_and_regulations,
               6: labor_law_enforcement,
               7: criminal_law_enforcement,
               8: government_actions,
               9: deliberative_data,
               10: goods_list}
    if ws_idx in options:
        if str(row[0]) == year:
            options[ws_idx](country, row)


def get_countries_key(elem):
    return elem.findtext("Name")


def get_goods_key(elem):
    return elem.findtext("Good_Name")


def indent(elem, level=0):
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


sanitize_re = re.compile(r"\(|\)|,|the")


def sanitize(str):
    str = re.sub(sanitize_re, "", str)
    str = str.replace("  ", " ")
    str = str.lower().replace(" ", "-")
    return str


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


skip = ["Instructions", "2.Overview of Children’s "]

for idx, sheet in enumerate(wb.sheetnames):
    if sheet in skip:
        continue
    ws = wb[sheet]
    for row in ws.iter_rows(min_row=2, values_only=True):
        country_name = row[1]
        country = country_exists(country_name)
        if country == None and country_name:
            country = ET.SubElement(countries, "Country")
            name = ET.SubElement(country, "Name")
            name.text = country_name
            ET.SubElement(
                country, "Webpage").text = "https://www.dol.gov/agencies/ilab/resources/reports/child-labor/" + sanitize(country_name)

        read_row(country, row, idx)


countries[:] = sorted(countries, key=get_countries_key)
indent(countries)
countries_tree.write(COUNTRIES_OUTPUT_FILE)

goods[:] = sorted(goods, key=get_goods_key)
indent(goods)
goods_tree.write(GOODS_OUTPUT_FILE)
