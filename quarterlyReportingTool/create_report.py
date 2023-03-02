import json
import gspread
import webcolors

import pandas as pd

from gspread_formatting import set_column_width
from gspread_dataframe import set_with_dataframe
from jira import JIRA, exceptions, resources
from typing import List, Optional
from quarterlyReportingTool.create_charts import create_pie_chart, create_bar_chart

configuration = json.load(open('resources/configuration.json'))


def create_report(team, quarter, planned_fte, planned_sp) -> str:
    planned_values = validate_input(team, quarter, planned_fte + planned_sp)
    if not planned_values:
        return '-1'

    jira = authorisation()
    if not jira:
        return '-1'

    return create_gspread(jira, team, quarter, planned_fte, planned_sp)


def validate_input(team, quarter, planned) -> List[float]:
    if team not in configuration['teams']:
        print("You have entered invalid team name. Try one from the following list: " + str(configuration['teams']))
        return []

    quarters = configuration['quarters'].keys()
    if quarter not in quarters:
        print("You have entered invalid quarter. Try one from the following list: " + str(quarters))
        return []

    planned_values = []
    for i in planned:
        try:
            planned_values.append(float(i))
        except ValueError:
            print("Invalid value for planned metrics on place " + str(i))
            return []

    return planned_values


def authorisation() -> Optional[JIRA]:
    token = configuration['token']
    jira = None
    try:
        jira = JIRA(server="https://issues.redhat.com/", token_auth=token)
        jira.myself()
    except exceptions.JIRAError:
        jira = None
        print("Invalid token")
    finally:
        return jira


def get_wp(jira: JIRA, team: str, quarter: str) -> Optional[List[resources.Issue]]:
    jql_request = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) AND " \
                  f"resolved >= '{configuration['quarters'][quarter][0]}'" \
                  f"AND resolved < '{configuration['quarters'][quarter][1]}' AND 'Epic Link' is not EMPTY " \
                  f"AND 'Story Points' is not EMPTY"
    return jira.search_issues(jql_request, maxResults=10000)


def get_release_operations(jira: JIRA, team: str, quarter: str) -> Optional[List[resources.Issue]]:
    jql_request1 = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) AND " \
                   f"EXD-WorkType = 'Release Operations' AND EXD-WorkType not in ('Maintenance', 'Technical " \
                   f"Improvement') AND resolved >= '{configuration['quarters'][quarter][0]}' AND resolved < '" \
                   f"{configuration['quarters'][quarter][1]}' AND 'Epic Link' is EMPTY AND 'Story Points' is not EMPTY"

    jql_request2 = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) " \
                   f"AND resolved >= '{configuration['quarters'][quarter][0]}' AND resolved < " \
                   f"'{configuration['quarters'][quarter][1]}' AND 'Parent Link' is not EMPTY AND 'Story Points' is " \
                   f"not EMPTY"
    return jira.search_issues(jql_request1, maxResults=10000) + jira.search_issues(jql_request2, maxResults=10000)


def get_maintenance(jira: JIRA, team: str, quarter: str) -> Optional[List[resources.Issue]]:
    jql_request = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) AND " \
                  f"EXD-WorkType not in ('Release Operations', 'Technical Improvement') AND EXD-WorkType = " \
                  f"'Maintenance' AND resolved >= '{configuration['quarters'][quarter][0]}' AND resolved < '" \
                  f"{configuration['quarters'][quarter][1]}' AND 'Epic Link' is EMPTY AND 'Story Points' is not EMPTY"
    return jira.search_issues(jql_request, maxResults=10000)


def get_standalone(jira: JIRA, team: str, quarter: str) -> Optional[List[resources.Issue]]:
    jql_request = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) AND " \
                  f"(EXD-WorkType not in ('Release Operations', 'Technical Improvement', 'Maintenance') " \
                  f"OR EXD-WorkType is EMPTY) AND resolved >= '{configuration['quarters'][quarter][0]}' " \
                  f"AND resolved < '{configuration['quarters'][quarter][1]}' AND 'Epic Link' is EMPTY " \
                  f"AND 'Story Points' is not EMPTY"
    return jira.search_issues(jql_request, maxResults=10000)


def get_issues_with_multiple_work_type(jira: JIRA, team: str, quarter: str) -> Optional[List[resources.Issue]]:
    jql_request = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) AND " \
                  f"((EXD-WorkType in ('Release Operations', 'Technical Improvement') AND EXD-WorkType = " \
                  f"'Maintenance') OR (EXD-WorkType in ('Maintenance', 'Technical Improvement') AND EXD-WorkType = " \
                  f"'Release Operations') OR " \
                  f"(EXD-WorkType in ('Maintenance', 'Release Operations') AND EXD-WorkType = 'Technical " \
                  f"Improvement')) AND resolved >= '{configuration['quarters'][quarter][0]}' AND resolved < '" \
                  f"{configuration['quarters'][quarter][1]}' " \
                  f"AND 'Story Points' is not EMPTY"
    return jira.search_issues(jql_request, maxResults=10000)


def get_issues_without_story_points(jira: JIRA, team: str, quarter: str) -> Optional[List[resources.Issue]]:
    jql_request = f"project={team} AND issuetype not in (Ticket, Sub-task, Epic) AND 'Story Points' is EMPTY " \
                  f"AND resolved >= '{configuration['quarters'][quarter][0]}' AND resolved <" \
                  f" '{configuration['quarters'][quarter][1]}' AND resolution not in " \
                  + "(" + '"Can\'t Do"' + ", 'Cannot Reproduce', Duplicate, 'Duplicate Ticket', 'Not a Bug', " \
                                          "Obsolete, Unresolved, " + '"Won\'t Do"' + ")"
    return jira.search_issues(jql_request, maxResults=10000)


def get_story_points(issues: List[resources.Issue]) -> int:
    sp = 0
    for issue in issues:
        sp += int(issue.get_field(configuration['custom_fields']['story_points']))
    return sp


def count_ratios(ftes: List[float], sps: List[float]) -> List[float]:
    ratio = []
    total_sps = sum(sps)
    for i in range(4):
        if ftes:
            ratio.append(round(ftes[i] / sum(ftes), 2))
        else:
            ratio.append(round(sps[i] / total_sps, 2))
    return ratio


def count_final_fte(planned_ftes: List[float], ratio: List[float]) -> List[float]:
    ftes = []
    planned_total_ftes = sum(planned_ftes)
    for i in range(4):
        ftes.append(ratio[i] * planned_total_ftes)
    return ftes


def create_google_sheet(team: str, quarter: str) -> gspread.Spreadsheet:
    service_account = gspread.service_account(filename="resources/service_account.json")
    # TODO handle log in error
    sheet = service_account.create(quarter + ' ' + team, configuration["report_path"])
    sheet.add_worksheet("Report", rows=100, cols=20)
    sheet.del_worksheet(sheet.worksheet("Sheet1"))
    return sheet


def create_error_reports(sheet: gspread.Spreadsheet, issues, name):
    if issues:
        df = pd.DataFrame()
        df["Key"] = list(map(lambda x: f'=HYPERLINK("https://issues.redhat.com/browse/{str(x)}", '
                                       f'"{str(x)}")', issues))
        df["Issue name"] = list(map(lambda x: x.fields.summary, issues))
        df["Status"] = list(map(lambda x: x.fields.status, issues))
        df["Created (GMT+0)"] = list(map(lambda x: x.fields.created[:10] + ' ' + x.fields.created[11:19], issues))
        df["Reporter"] = list(map(lambda x: x.fields.reporter, issues))
        df["Assignee"] = list(map(lambda x: x.fields.assignee, issues))
        set_with_dataframe(sheet.add_worksheet(name, rows=10000, cols=1), df)


def create_gspread(jira: JIRA, team: str, quarter: str, planned_ftes: List[float], planned_sps: List[float]) -> str:
    sheet = create_google_sheet(team, quarter)

    no_sp_issues = get_issues_without_story_points(jira, team, quarter)
    create_error_reports(sheet, no_sp_issues, "Issues without story points")

    multiple_exd_issues = get_issues_with_multiple_work_type(jira, team, quarter)
    create_error_reports(sheet, multiple_exd_issues, "Issues with multiple EXD-WorkType")

    format_sheet(sheet)

    work_packages = get_wp(jira, team, quarter)
    release_operations = get_release_operations(jira, team, quarter)
    maintenance = get_maintenance(jira, team, quarter)
    standalone = get_standalone(jira, team, quarter)

    planned_ratio = count_ratios(planned_ftes, planned_sps)
    final_sps = [get_story_points(work_packages), get_story_points(release_operations),
                 get_story_points(maintenance), get_story_points(standalone)]
    final_ratio = count_ratios([], final_sps)
    final_ftes = count_final_fte(planned_ftes, final_ratio)
    same_as_planned = ['', 'Same as planned', 'Same as planned']
    user_input = ['', 'User input', 'User input']
    no_sp = ['', 'This is not captured by the Story points', '']

    df = pd.DataFrame(index=range(7))
    df["Total Available capacity"] = ['Change Portfolio', 'Business As Usual', '', '', '', '', '']
    df[str(sum(planned_ftes))] = list(configuration['metrics_names'])
    df["Planned FTEs"] = [str(i) for i in planned_ftes] + user_input
    df["Planned SPs"] = [str(i) for i in planned_sps] + no_sp
    df["Planned capacity distribution"] = [str(i) for i in planned_ratio] + user_input
    df["Final FTEs"] = [str(i) for i in final_ftes] + same_as_planned
    df["Final SPs"] = [str(i) for i in final_sps] + no_sp
    df["Final capacity distribution"] = [str(i) for i in final_ratio] + same_as_planned
    df["Diff planned vs real"] = ['1', '-2', '3', '-4'] + same_as_planned

    set_with_dataframe(sheet.worksheet("Report"), df)

    create_pie_chart(sheet.id, sheet.worksheet("Report").id, 'Planned Capacity', [1, 2, 5, 3], [10, 0])
    create_pie_chart(sheet.id, sheet.worksheet("Report").id, 'Actual Capacity', [1, 5, 5, 6], [10, 2])
    create_bar_chart(sheet.id, sheet.worksheet("Report").id)

    return sheet.url


def format_sheet(sheet: gspread.Spreadsheet) -> None:
    report_wsh = sheet.worksheet("Report")
    report_wsh.merge_cells('D7:D8')
    report_wsh.merge_cells('G7:G8')
    report_wsh.merge_cells('A3:A8')
    set_column_width(report_wsh, 'A:I', 245)

    dark_grey = webcolors.hex_to_rgb(configuration['colors']['dark_grey'])
    cell_style_dark_grey = {
        "backgroundColor": {
            "red": dark_grey[0] / 255,
            "green": dark_grey[1] / 255,
            "blue": dark_grey[2] / 255
        },
        "textFormat": {
            "fontSize": 13,
            "bold": True
        }}

    grey = webcolors.hex_to_rgb(configuration['colors']['grey'])
    cell_style_grey = {
        "backgroundColor": {
            "red": grey[0] / 255,
            "green": grey[1] / 255,
            "blue": grey[2] / 255
        }}

    light_grey = webcolors.hex_to_rgb(configuration['colors']['light_grey'])
    cell_style_light_grey = {
        "backgroundColor": {
            "red": light_grey[0] / 255,
            "green": light_grey[1] / 255,
            "blue": light_grey[2] / 255
        }}

    orange = webcolors.hex_to_rgb(configuration['colors']['orange'])
    cell_style_orange = {
        "backgroundColor": {
            "red": orange[0] / 255,
            "green": orange[1] / 255,
            "blue": orange[2] / 255
        }}

    light_orange = webcolors.hex_to_rgb(configuration['colors']['light_orange'])
    cell_style_light_orange = {
        "backgroundColor": {
            "red": light_orange[0] / 255,
            "green": light_orange[1] / 255,
            "blue": light_orange[2] / 255
        }}

    red = webcolors.hex_to_rgb(configuration['colors']['red'])
    cell_style_red = {
        "backgroundColor": {
            "red": red[0] / 255,
            "green": red[1] / 255,
            "blue": red[2] / 255
        }}

    light_red = webcolors.hex_to_rgb(configuration['colors']['light_red'])
    cell_style_light_red = {
        "backgroundColor": {
            "red": light_red[0] / 255,
            "green": light_red[1] / 255,
            "blue": light_red[2] / 255
        }}

    properties = {
        "borders": {
            "top": {"style": "SOLID"},
            "bottom": {"style": "SOLID"},
            "left": {"style": "SOLID"},
            "right": {"style": "SOLID"}
        },
        "horizontalAlignment": "CENTER",
        "verticalAlignment": "middle",
    }

    report_wsh.format("A1:I8", properties)
    report_wsh.format("A1:A8", cell_style_dark_grey)
    report_wsh.format("B2:B8", cell_style_dark_grey)
    report_wsh.format("C1:I1", cell_style_dark_grey)
    report_wsh.format("C2:C8", cell_style_light_orange)
    report_wsh.format("B1", cell_style_light_orange)
    report_wsh.format("E2:E8", cell_style_orange)
    report_wsh.format("F2:H6", cell_style_light_grey)
    report_wsh.format("F7:H8", cell_style_grey)
    report_wsh.format("I2:I6", cell_style_light_red)
    report_wsh.format("I7:I8", cell_style_red)

    try:
        no_sp = sheet.worksheet("Issues without story points")
        set_column_width(no_sp, 'A:F', 200)
        set_column_width(no_sp, 'B', 700)
        no_sp.format("A1:F1", cell_style_dark_grey)
        no_sp.format("A:F", properties)
    except gspread.exceptions.WorksheetNotFound:
        pass

    try:
        exd = sheet.worksheet("Issues with multiple EXD-WorkType")
        set_column_width(exd, 'A:F', 200)
        set_column_width(exd, 'B', 700)
        exd.format("A1:F1", cell_style_dark_grey)
        exd.format("A:F", properties)
    except gspread.exceptions.WorksheetNotFound:
        pass
