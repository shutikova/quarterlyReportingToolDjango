import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from typing import List

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


def create_service():
    service = None
    credentials = None

    if os.path.exists('resources/token.json'):
        credentials = Credentials.from_authorized_user_file('resources/token.json', SCOPES)
    # If there are no credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'resources/credentials.json', SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('resources/token.json', 'w') as token:
            token.write(credentials.to_json())

    try:
        service = build('sheets', 'v4', credentials=credentials)
    except HttpError as err:
        print(err)
    finally:
        return service


def create_pie_chart(spreadsheet_id: str, worksheet_id: str, name: str, cells: List[int], position: List[int]) -> None:
    service = create_service()

    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={
        "requests": [
            {
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": name,
                            "pieChart": {
                                "legendPosition": "BOTTOM_LEGEND",
                                "domain": {
                                    "sourceRange": {
                                        "sources": [
                                            {
                                                "sheetId": worksheet_id,
                                                "startRowIndex": 1,
                                                "startColumnIndex": 1,
                                                "endRowIndex": 5,
                                                "endColumnIndex": 2
                                            }
                                        ]
                                    }
                                },
                                "series": {
                                    "sourceRange": {
                                        "sources": [
                                            {
                                                "sheetId": worksheet_id,
                                                "startRowIndex": cells[0],
                                                "startColumnIndex": cells[1],
                                                "endRowIndex": cells[2],
                                                "endColumnIndex": cells[3]
                                            }
                                        ]
                                    }
                                },
                                "pieHole": 0.4
                            }
                        },
                        "position": {
                            "overlayPosition": {
                                "anchorCell": {
                                    "sheetId": worksheet_id,
                                    "rowIndex": position[0],
                                    "columnIndex": position[1]
                                },
                                "widthPixels": 400,
                                "heightPixels": 300
                            }
                        }
                    }
                }
            }
        ]
    })
    request.execute()


def create_bar_chart(spreadsheet_id: str, worksheet_id: str):
    service = create_service()
    request = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={
        "requests": [
            {
                "addChart": {
                    "chart": {
                        "spec": {
                            "title": "Reality compared to the plan",
                            "basicChart": {
                                "chartType": "COLUMN",
                                "legendPosition": "BOTTOM_LEGEND",
                                "domains": [
                                    {
                                        "domain": {
                                            "sourceRange": {
                                                "sources": [
                                                    {
                                                        "sheetId": worksheet_id,
                                                        "startRowIndex": 0,
                                                        "endRowIndex": 8,
                                                        "startColumnIndex": 1,
                                                        "endColumnIndex": 2
                                                    }
                                                ]
                                            }
                                        }
                                    }
                                ],
                                "series": [
                                    {
                                        "series": {
                                            "sourceRange": {
                                                "sources": [
                                                    {
                                                        "sheetId": worksheet_id,
                                                        "startRowIndex": 0,
                                                        "endRowIndex": 6,
                                                        "startColumnIndex": 8,
                                                        "endColumnIndex": 9
                                                    }
                                                ]
                                            },
                                        },
                                    }
                                ],
                                "headerCount": 1
                            }
                        },
                        "position": {
                            "overlayPosition": {
                                "anchorCell": {
                                    "sheetId": worksheet_id,
                                    "rowIndex": 10,
                                    "columnIndex": 4
                                },
                                "offsetXPixels": 0,
                                "offsetYPixels": 0,
                                "widthPixels": 900,
                                "heightPixels": 500
                            }
                        }
                    }
                }
            }
        ]

    })

    request.execute()
