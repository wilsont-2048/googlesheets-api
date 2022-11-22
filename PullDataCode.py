# Copyright 2018 Google LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# [START sheets_quickstart]
from __future__ import print_function
from cgi import test
from email import message

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
 
# Spreadsheet ID relating to the spreadsheet that contains endorsement info.
SPREADSHEET_ID = '1F1bM6jpW2_07Yys5UzRZr3rnQ75P16emwfehf1Xy9tM'
# SAMPLE_RANGE_NAME = 'Class Data!A2:E'

def get_spreadsheet_data(range_name):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=range_name).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        return values
    except HttpError as err:
        print(err)


def main():
    org_values = get_spreadsheet_data('Organizations!A2:D')
    feds_values = get_spreadsheet_data('Federal Staff!A2:J')
    state_county_values = get_spreadsheet_data('State and County Staff!A2:F')
    city_values = get_spreadsheet_data('City Staff!A2:F')
    ed_values = get_spreadsheet_data('Education Staff!A2:F')

    writehtml(org_values, feds_values, state_county_values, city_values, ed_values)


def writehtml(org, feds, stct, city, ed):
    EnWriteFunc = open("PulledDataPage_en.html", "w")
    ViWriteFunc = open("PulledDataPage_vi.html", "w")
    EsWriteFunc = open("PulledDataPage_es.html", "w")

    write_org(org, EnWriteFunc, ViWriteFunc, EsWriteFunc)
    write_feds(feds, EnWriteFunc, ViWriteFunc, EsWriteFunc)
    write_state_county(stct, EnWriteFunc, ViWriteFunc, EsWriteFunc)
    write_city(city, EnWriteFunc, ViWriteFunc, EsWriteFunc)
    write_ed(ed, EnWriteFunc, ViWriteFunc, EsWriteFunc)

    EnWriteFunc.close()
    ViWriteFunc.close()
    EsWriteFunc.close()


def write_org(data, en_func, vi_func, es_func):
    orgdetailsbegin = """        <details>\n"""

    en_func.write(orgdetailsbegin)
    vi_func.write(orgdetailsbegin)
    es_func.write(orgdetailsbegin)

    en_orgbegin = """            <summary>Organizations</summary>
            <section class="grid-profiles">
                <div class="squeeze">
                    <div class="modified-wrap">
                        <div class="guts kinda-full">
                            <div class="columns limit with-max with-margin to-4 then-3 finally-2">"""   

    vi_orgbegin = """            <summary>Các tổ chức</summary>
            <section class="grid-profiles">
                <div class="squeeze">
                    <div class="modified-wrap">
                        <div class="guts kinda-full">
                            <div class="columns limit with-max with-margin to-4 then-3 finally-2">"""   

    es_orgbegin = """            <summary>Organizaciones</summary>
            <section class="grid-profiles">
                <div class="squeeze">
                    <div class="modified-wrap">
                        <div class="guts kinda-full">
                            <div class="columns limit with-max with-margin to-4 then-3 finally-2">"""   

    en_func.write(en_orgbegin)
    vi_func.write(vi_orgbegin)
    es_func.write(es_orgbegin)

    print('Organizations')
    print('Name, Image Link')
    for row in data:
        print('%s, %s' % (row[0], row[1]))
        message = f"""\n                                <div class="column item">
                                    <div class="text-stuff">
                                        <div class="top has-image full-style">
                                            <img src="{row[1]}" alt="{row[0]}" width="400" height="400" class="alignnone size-full wp-image-615" />
                                            <div class="words-modified">
                                                <h5 class="card-title">{row[0]}</h5>
                                            </div>
                                        </div>
                                    </div>
                                </div>"""
        en_func.write(message)
        vi_func.write(message)
        es_func.write(message)

    orgend = """\n                            </div>
                        </div>
                    </div>
                </div>
            </section>
        </details>"""

    en_func.write(orgend)
    vi_func.write(orgend)
    es_func.write(orgend)


def write_feds(data, en_func, vi_func, es_func):
    fedsdetailsbegin = """\n        <details>\n"""

    en_func.write(fedsdetailsbegin)
    vi_func.write(fedsdetailsbegin)
    es_func.write(fedsdetailsbegin)

    en_fedsbegin = """            <summary>Federal Staff</summary>
            <section class="grid-profiles">
                <div class="squeeze">
                    <div class="modified-wrap">
                        <div class="guts kinda-full">
                            <div class="columns limit with-max with-margin to-4 then-3 finally-2">"""
    
    vi_fedsbegin = """            <summary>Nhân viên liên bang</summary>
            <section class="grid-profiles">
                <div class="squeeze">
                    <div class="modified-wrap">
                        <div class="guts kinda-full">
                            <div class="columns limit with-max with-margin to-4 then-3 finally-2">"""
    
    es_fedsbegin = """            <summary>Personal Federal</summary>
            <section class="grid-profiles">
                <div class="squeeze">
                    <div class="modified-wrap">
                        <div class="guts kinda-full">
                            <div class="columns limit with-max with-margin to-4 then-3 finally-2">"""

    en_func.write(en_fedsbegin)
    vi_func.write(vi_fedsbegin)
    es_func.write(es_fedsbegin)

    print('Federal Staff')
    print('Name, Location, Image Link')
    for row in data:
        print('%s, %s, %s' % (row[0], row[2], row[5]))
        message = f"""\n                                <div class="column item">
                                    <div class="text-stuff">
                                        <div class="top has-image full-style">
                                            <img src="{row[5]}" alt="{row[0]}" width="400" height="400" class="alignnone size-full wp-image-615" />
                                            <div class="words">
                                                <h3>{row[0]}</h3>
                                                <p class="superheader">{row[4]}</p>
                                            </div>
                                        </div>
                                    </div>
                                </div>"""
        en_func.write(message)
        vi_func.write(message)
        es_func.write(message)

    fedsend = """\n                            </div>
                        </div>
                    </div>
                </div>
            </section>
        </details>"""

    en_func.write(fedsend)
    vi_func.write(fedsend)
    es_func.write(fedsend)


def write_state_county(data, en_func, vi_func, es_func):
    stctdetailsbegin = """\n        <details>\n"""

    en_func.write(stctdetailsbegin)
    vi_func.write(stctdetailsbegin)
    es_func.write(stctdetailsbegin)

    en_stctbegin = """            <summary>State and County Staff</summary>
            <div class="list-container">
                <ul>"""

    vi_stctbegin = """            <summary>Nhân viên Tiểu bang và Quận</summary>
            <div class="list-container">
                <ul>"""

    es_stctbegin = """            <summary>Personal estatal y del condado</summary>
            <div class="list-container">
                <ul>"""

    en_func.write(en_stctbegin)
    vi_func.write(vi_stctbegin)
    es_func.write(es_stctbegin)

    print('State and County Staff')
    print('Name, Position, Category')
    for row in data:
        print('%s, %s, %s' % (row[0], row[1], row[2]))
        message = f"""\n                    <li>{row[1]} {row[0]}</li>"""
        en_func.write(message)
        vi_func.write(message)
        es_func.write(message)

    stctend = """\n                </ul> 
            </div>
        </details>"""

    en_func.write(stctend)
    vi_func.write(stctend)
    es_func.write(stctend)


def write_city(data, en_func, vi_func, es_func):
    citydetailsbegin = """\n        <details>\n"""

    en_func.write(citydetailsbegin)
    vi_func.write(citydetailsbegin)
    es_func.write(citydetailsbegin)

    en_citybegin = """            <summary>City Staff</summary>
            <div class="list-container">
                <ul>"""
    vi_citybegin = """            <summary>Nhân viên thành phố</summary>
            <div class="list-container">
                <ul>"""
    es_citybegin = """            <summary>Personal de la Ciudad</summary>
            <div class="list-container">
                <ul>"""

    en_func.write(en_citybegin)
    vi_func.write(vi_citybegin)
    es_func.write(es_citybegin)

    print('City Staff')
    print('Name, Position, Category')
    for row in data:
        print('%s, %s, %s' % (row[0], row[1], row[2]))
        message = f"""\n                    <li>{row[1]} {row[0]}</li>"""
        en_func.write(message)
        vi_func.write(message)
        es_func.write(message)

    cityend = """\n                </ul>
            </div>
        </details>"""

    en_func.write(cityend)
    vi_func.write(cityend)
    es_func.write(cityend)


def write_ed(data, en_func, vi_func, es_func):
    eddetailsbegin = """\n        <details>\n"""

    en_func.write(eddetailsbegin)
    vi_func.write(eddetailsbegin)
    es_func.write(eddetailsbegin)

    en_edbegin = """            <summary>Education Staff</summary>
            <div class="list-container">
                <ul>"""

    vi_edbegin = """            <summary>Nhân viên giáo dục</summary>
            <div class="list-container">
                <ul>"""

    es_edbegin = """            <summary>personal educativo</summary>
            <div class="list-container">
                <ul>"""

    en_func.write(en_edbegin)
    vi_func.write(vi_edbegin)
    es_func.write(es_edbegin)

    print('Education Staff')
    print('Name, Position')
    for row in data:
        if row[4] == "Pending":
            print(print('Status \"Pending\":' + '%s, %s' % (row[0], row[1])))
        else:
            print('%s, %s' % (row[0], row[1]))
            message = f"""\n                    <li>{row[1]} {row[0]}</li>"""
            en_func.write(message)
            vi_func.write(message)
            es_func.write(message)

    edend = """\n                </ul>
            </div>
        </details>"""

    en_func.write(edend)
    vi_func.write(edend)
    es_func.write(edend)


if __name__ == '__main__':
    main()
# [END sheets_quickstart]