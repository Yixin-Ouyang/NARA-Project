import pandas as pd
import tkinter as tk
from tkinter import filedialog

# pop up a window and ask user to select xlsx file
root = tk.Tk()
root.withdraw()

# get file path and name
file_path = filedialog.askopenfilename()
output_path = file_path[:-5]

# get the data from the spreadsheet
df = pd.read_excel(file_path, sheet_name='FileUnitTemplate', dtype=object)
checklist = pd.read_excel(file_path, sheet_name='Lists', dtype=object)

# strip the header
df = df.rename(columns=lambda x: x.strip())
checklist = checklist.rename(columns=lambda x: x.strip())

# get the blank table
blank = df.isnull()

print('Creating DAS XML...')

# create xml file
xmlData = open(output_path + '.xml', 'w', encoding='utf-8')
xmlData.write('<?xml version="1.0" encoding="utf-8"?>\n')
xmlData.write('<import xmlns="http://ui.das.nara.gov/">\n\t<fileUnitArray>')

# loop each row in xlsx file
for i in df.index:
    # validation
    # template variables
    # required variables
    # check data_control_group
    if blank['dataControlGroup'][i]:
        data_control_group = '[DATA CONTROL GROUP REQUIRED]'
        reference_unit = '[REFERENCE UNIT REQUIRED]'
        location = '[LOCATION REQUIRED]'
        ou_group = '[OU GROUP REQUIRED]'
        print('Blank dataControlGroup found in line ', i+2)
    else:
        data_control_group = df['dataControlGroup'][i]
        if data_control_group == 'LL':
            reference_unit = 'Center for Legislative Archives'
            location = 'National Archives Building - Archives I (Washington, DC)'
            ou_group = 'NWL'
        elif data_control_group == 'LM':
            reference_unit = 'Presidential Materials Division'
            location = 'National Archives Building - Archives I (Washington, DC)'
            ou_group = 'NLMS'
        elif data_control_group == 'LPBHO':
            reference_unit = 'Barack Obama Presidential Library'
            location = 'Barack Obama Presidential Library (Hoffman Estates, IL)'
            ou_group = 'LPBHO'
        elif data_control_group == 'LPDDE':
            reference_unit = 'Dwight D. Eisenhower Library'
            location = 'Dwight D. Eisenhower Library (Abilene, KS)'
            ou_group = 'NLDDE'
        elif data_control_group == 'LPFDR':
            reference_unit = 'Franklin D. Roosevelt Library'
            location = 'Franklin D. Roosevelt Library (Hyde Park, NY)'
            ou_group = 'NLFDR'
        elif data_control_group == 'LPGB':
            reference_unit = 'George Bush Library'
            location = 'George Bush Library (College Station, TX)'
            ou_group = 'NLGB'
        elif data_control_group == 'LPGWB':
            reference_unit = 'George W. Bush Library'
            location = 'George W. Bush Library (Lewisville, TX)'
            ou_group = 'NLGWB'
        elif data_control_group == 'LPGRF':
            reference_unit = 'Gerald R. Ford Library'
            location = 'Gerald R. Ford Library (Ann Arbor, MI)'
            ou_group = 'NLGRF'
        elif data_control_group == 'LPHH':
            reference_unit = 'Herbert Hoover Library'
            location = 'Herbert Hoover Library (West Branch, IA)'
            ou_group = 'NLHH'
        elif data_control_group == 'LPHST':
            reference_unit = 'Harry S. Truman Library'
            location = 'Harry S. Truman Library (Independence, MO)'
            ou_group = 'NLHST'
        elif data_control_group == 'LPJC':
            reference_unit = 'Jimmy Carter Library'
            location = 'Jimmy Carter Library (Atlanta, GA)'
            ou_group = 'NLJC'
        elif data_control_group == 'LPJFK':
            reference_unit = 'John F. Kennedy Library'
            location = 'John F. Kennedy Library (Boston, MA)'
            ou_group = 'NLJFK'
        elif data_control_group == 'LPLBJ':
            reference_unit = 'Lyndon B. Johnson Library'
            location = 'Lyndon Baines Johnson Library (Austin, TX)'
            ou_group = 'NLLBJ'
        elif data_control_group == 'LPRN':
            reference_unit = 'Richard Nixon Library'
            location = 'Richard Nixon Library (Yorba Linda, CA)'
            ou_group = 'NLRN'
        elif data_control_group == 'LPRR':
            reference_unit = 'Ronald Reagan Library'
            location = 'Ronald Reagan Library (Simi Valley, CA)'
            ou_group = 'NLRR'
        elif data_control_group == 'LPWJC':
            reference_unit = 'William J. Clinton Library'
            location = 'William J. Clinton Library (Little Rock, AR)'
            ou_group = 'NLWJC'
        elif data_control_group == 'RDF':
            reference_unit = 'National Archives at College Park - FOIA'
            location = 'National Archives at College Park - Archives II (College Park, MD)'
            ou_group = 'RDF'
        elif data_control_group == 'RDSC':
            reference_unit = 'National Archives at College Park - Cartographic'
            location = 'National Archives at College Park - Archives II (College Park, MD)'
            ou_group = 'NWCS-C'
        elif data_control_group == 'RDSM':
            reference_unit = 'National Archives at College Park - Motion Pictures'
            location = 'National Archives at College Park - Archives II (College Park, MD)'
            ou_group = 'NWCS-M'
        elif data_control_group == 'RDSS':
            reference_unit = 'National Archives at College Park - Still Pictures'
            location = 'National Archives at College Park - Archives II (College Park, MD)'
            ou_group = 'NWCS-S'
        elif data_control_group == 'RDTP1':
            reference_unit = 'National Archives at Washington, DC - Textual Reference'
            location = 'National Archives Building - Archives I (Washington, DC)'
            ou_group = 'RDTP1'
        elif data_control_group == 'RDTP2':
            reference_unit = 'National Archives at College Park, MD - Textual Reference'
            location = 'National Archives at College Park - Archives II (College Park, MD)'
            ou_group = 'RDTP2'
        elif data_control_group == 'RDEP':
            reference_unit = 'National Archives at College Park - Electronic Records'
            location = 'National Archives at College Park - Archives II (College Park, MD)'
            ou_group = 'NWME'
        elif data_control_group == 'REAT':
            reference_unit = 'National Archives at Atlanta'
            location = 'NARA\'s Southeast Region (Atlanta, GA)'
            ou_group = 'NRCA'
        elif data_control_group == 'REBO':
            reference_unit = 'National Archives at Boston'
            location = 'NARA\'s Northeast Region (Boston, MA)'
            ou_group = 'NRAAB'
        elif data_control_group == 'RENY':
            reference_unit = 'National Archives at New York'
            location = 'NARA\'s Northeast Region (New York City, NY)'
            ou_group = 'NRAAN'
        elif data_control_group == 'REPA':
            reference_unit = 'National Archives at Philadelphia'
            location = 'NARA\'s Mid Atlantic Region (Philadelphia, PA)'
            ou_group = 'NRBA'
        elif data_control_group == 'RLSL':
            reference_unit = 'National Personnel Records Center - Military Personnel Records'
            location = 'National Military Personnel Records Center (St. Louis, MO)'
            ou_group = 'NRPA'
        elif data_control_group == 'RMCH':
            reference_unit = 'National Archives at Chicago'
            location = 'NARA\'s Great Lakes Region (Chicago, IL)'
            ou_group = 'NRDA'
        elif data_control_group == 'RMDV':
            reference_unit = 'National Archives at Denver'
            location = 'NARA\'s Rocky Mountain Region (Denver, CO)'
            ou_group = 'NRGA'
        elif data_control_group == 'RMFW':
            reference_unit = 'National Archives at Fort Worth'
            location = 'NARA\'s Southwest Region (Fort Worth, TX)'
            ou_group = 'NRFA'
        elif data_control_group == 'RMKC':
            reference_unit = 'National Archives at Kansas City'
            location = 'NARA\'s Central Plains Region (Kansas City, MO)'
            ou_group = 'NREA'
        elif data_control_group == 'RWRS':
            reference_unit = 'National Archives at Riverside'
            location = 'NARA\'s Pacific Region (Riverside, CA)'
            ou_group = 'NRHAR'
        elif data_control_group == 'RWSB':
            reference_unit = 'National Archives at San Francisco'
            location = 'NARA\'s Pacific Region (San Bruno, CA)'
            ou_group = 'NRHAS'
        elif data_control_group == 'RWSE':
            reference_unit = 'National Archives at Seattle'
            location = 'NARA\'s Pacific Alaska Region (Seattle, WA)'
            ou_group = 'NRIAS'
        else:
            data_control_group = '[DATA CONTROL GROUP REQUIRED]'
            reference_unit = '[REFERENCE UNIT REQUIRED]'
            location = '[LOCATION REQUIRED]'
            ou_group = '[OU GROUP REQUIRED]'
            print('Undefined dataControlGroup found in line ', i+2)

    # check parent series naid
    if blank['parentSeriesNaid'][i]:
        parent_series_naid = '[PARENT SERIES NAID REQUIRED]'
        print('Blank parentSeriesNaid found in line ', i+2)
    else:
        parent_series_naid = df['parentSeriesNaid'][i]
        if not isinstance(parent_series_naid, int):
            parent_series_naid = '[PARENT SERIES NAID REQUIRED]'
            print('Undefined parentSeriesNaid found in line ', i+2)

    # check title
    if blank['title'][i]:
        title = '[TITLE REQUIRED]'
        print('Blank title found in line ', i+2)
    else:
        title = df['title'][i]

    # check access restriction status
    if blank['accessRestrictionStatus'][i]:
        access_restriction_status = '[ACCESS RESTRICTION STATUS REQUIRED]'
        print('Blank accessRestrictionStatus found in line ', i+2)
    else:
        access_restriction_status = df['accessRestrictionStatus'][i]
        if access_restriction_status not in checklist['Access Restriction Status'].tolist():
            access_restriction_status = '[ACCESS RESTRICTION STATUS REQUIRED]'
            print('Undefined accessRestrictionStatus found in line ', i+2)

    # check use restriction status
    if blank['useRestrictionStatus'][i]:
        use_restriction_status = '[USE RESTRICTION STATUS REQUIRED]'
        print('Blank useRestrictionStatus found in line ', i+2)
    else:
        use_restriction_status = df['useRestrictionStatus'][i]
        if use_restriction_status not in checklist['Use Restriction Status'].tolist():
            use_restriction_status = '[USE RESTRICTION STATUS REQUIRED]'
            print('Undefined useRestrictionStatus found in line ', i+2)

    # check general record type
    if blank['generalRecordsType'][i]:
        general_records_type = '[GENERAL RECORDS TYPE REQUIRED]'
        print('Blank generalRecordsType found in line ', i+2)
    else:
        general_records_type = df['generalRecordsType'][i]
        if general_records_type not in checklist['General Records Type'].tolist():
            general_records_type = '[GENERAL RECORDS TYPE REQUIRED]'
            print('Undefined generalRecordsType found in line ', i+2)

    # check copy status
    if blank['copyStatus'][i]:
        copy_status = '[COPY STATUS REQUIRED]'
        print('Blank copyStatus found in line ', i+2)
    else:
        copy_status = df['copyStatus'][i]
        if copy_status not in checklist['Copy Status'].tolist():
            copy_status = '[COPY STATUS REQUIRED]'
            print('Undefined copyStatus found in line ', i+2)

    # check specific media type
    if blank['specificMediaType'][i]:
        specific_media_type = '[SPECIFIC MEDIA TYPE REQUIRED]'
        print('Blank specificMediaType found in line ', i+2)
    else:
        specific_media_type = df['specificMediaType'][i]
        if specific_media_type not in checklist['Specific Media Type'].tolist():
            specific_media_type = '[SPECIFIC MEDIA TYPE REQUIRED]'
            print('Undefined specificMediaType found in line ', i+2)

    # check general media type
    if blank['generalMediaType'][i]:
        general_media_type = '[GENERAL MEDIA TYPE REQUIRED]'
        print('Blank generalMediaType found in line ', i+2)
    else:
        general_media_type = df['generalMediaType'][i]
        if general_media_type not in checklist['General Media Type'].tolist():
            general_media_type = '[GENERAL MEDIA TYPE REQUIRED]'
            print('Undefined generalMediaType found in line ', i+2)

    # optional variable
    # check container id
    if not blank['containerId'][i]:
        container_id = df['containerId'][i]
        if not isinstance(container_id, int):
            container_id = '[CONTAINER ID REQUIRED]'
            print('Undefined containerId found in line ', i+2)

    # check Image File Name(s) -- Ranges or Directory Name OK
    image_file_name = df['Image File Name(s) -- Ranges or Directory Name OK'][i]

    # check scope and content note
    scope_and_content_note = df['scopeAndContentNote'][i]

    # check specific access restriction
    specific_access_restriction = df['specificAccessRestriction'][i]
    if specific_access_restriction not in checklist['Specific Access Restriction'].tolist():
        specific_access_restriction = '[SPECIFIC ACCESS RESTRICTION REQUIRED]'
        print('Undefined specificAccessRestriction found in line ', i+2)

    # check security classification
    security_classification = df['securityClassification'][i]
    if security_classification not in checklist['Security Classification'].tolist():
        security_classification = '[SECURITY CLASSIFICATION REQUIRED]'
        print('Wrong securityClassification found in line ', i+2)

    # check access restriction note
    access_restriction_note = df['accessRestrictionNote'][i]

    # check specific use restriction
    specific_use_restriction = df['specificUseRestriction'][i]
    if specific_use_restriction not in checklist['Specific Use Restriction'].tolist():
        specific_use_restriction = '[SPECIFIC USE RESTRICTION REQUIRED]'
        print('Undefined specific restriction found in line ', i+2)

    # check use restriction note
    use_restriction_note = df['useRestrictionNote'][i]

    # check if there are additional columns
    extra_columns_exist = False
    try:
        variant_control_number_type = df['variantControlNumberType'][i]
        variant_control_number_num = df['variantControlNumberNum'][i]
        if not isinstance(variant_control_number_num, int):
            variant_control_number_num = '[VARIANT CONTROL NUMBER NUM REQUIRED]'
            print('Undefined variantControlNumberNum found in line ', i+2)
        variant_control_number_note = df['variantControlNumberNote'][i]
        extra_columns_exist = True
    except KeyError:
        variant_control_number_type = ""
        variant_control_number_num = ""
        variant_control_number_note = ""

    # write this file unit to xml file
    das_xml = """
        <fileUnit xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <sequenceOrder>""" + str(i+1) + """</sequenceOrder>
            <dataControlGroup>
                <groupCd>""" + str(data_control_group) + """</groupCd>
                <groupId>ou=""" + str(ou_group) + """,ou=groups</groupId>
            </dataControlGroup>
            <parentSeries>
                <naId>""" + str(parent_series_naid) + """</naId>
            </parentSeries>
            <localIdentifier/>
            <title>""" + str(title) + """</title>
            <dateNote/>
            <arrangement/>
            <scopeAndContentNote/>
            <generalRecordsTypeArray>
                <generalRecordsType>
                    <termName>""" + str(general_records_type) + """</termName>
                </generalRecordsType>
            </generalRecordsTypeArray>
            <physicalOccurrenceArray>
                <fileUnitPhysicalOccurrence>
                    <copyStatus>
                        <termName>""" + str(copy_status) + """</termName>
                    </copyStatus>
                    <locationArray>
                        <location>
                            <facility>
                                <termName>""" + str(location) + """</termName>
                            </facility>
                        </location>
                    </locationArray>
                    <totalRunningTime/>
                    <mediaOccurrenceArray>
                        <mediaOccurrence>
                            <specificMediaType>
                                <termName>""" + str(specific_media_type) + """</termName>
                            </specificMediaType>
                            <physicalRestrictionNote/>
                            <containerId/>
                            <generalMediaTypeArray>
                                <generalMediaType>
                                    <termName>""" + str(general_media_type) + """</termName>
                                </generalMediaType>
                            </generalMediaTypeArray>
                        </mediaOccurrence>
                    </mediaOccurrenceArray>
                    <referenceUnitArray>
                        <referenceUnit>
                            <termName>""" + str(reference_unit) + """</termName> 
                        </referenceUnit>
                    </referenceUnitArray>
                </fileUnitPhysicalOccurrence>
            </physicalOccurrenceArray>"""

    if extra_columns_exist:
        das_xml += """
            <variantControlNumberArray>
                <variantControlNumber>
                   <type>
                      <termName>""" + str(variant_control_number_type) + """</termName>
                   </type>
                   <number>""" + str(variant_control_number_num) + """</number>
                   <note>""" + str(variant_control_number_note) + """</note>
                </variantControlNumber>
             </variantControlNumberArray>"""

    das_xml += """
            <accessRestriction>
                <status>
                    <termName>""" + str(access_restriction_status) + """</termName>
                </status>"""

    if not blank['accessRestrictionNote'][i]:
        das_xml += """
                <accessRestrictionNote>""" + str(access_restriction_note) + """</accessRestrictionNote>"""

    if not blank['specificAccessRestriction'][i]:
        das_xml += """
                <specificAccessRestrictionArray>
                   <specificAccessRestriction>
                      <restriction>
                         <termName>""" + str(specific_access_restriction) + """</termName>
                      </restriction>
                   </specificAccessRestriction>
                </specificAccessRestrictionArray>"""

    das_xml += """
            </accessRestriction>
            <useRestriction>
                <note/>
                <status>
                    <termName>""" + str(use_restriction_status) + """</termName>
                </status>
            </useRestriction>
            <staffOnlyNote/>
        </fileUnit>"""

    xmlData.write(das_xml)

xmlData.write('\n\t</fileUnitArray>\n</import>')
print('DAS XML complete!')
xmlData.close()
