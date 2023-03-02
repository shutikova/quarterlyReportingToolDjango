from django.http import HttpResponse
from django.shortcuts import render

from quarterlyReportingTool.models import Quarter, Team
from quarterlyReportingTool.create_report import create_report


def index(request):
    context = {
        'quarters_list': Quarter.objects.order_by('-quarter_text'),
        'teams_list': Team.objects.order_by('-team_text'),
    }

    return render(request, 'quarterlyReportingTool/generate_report.html', context)


def results(request):
    url = create_report('RHELBLD', 'CY22Q1', [1, 2, 3, 4], [5, 6, 7, 8])
    return HttpResponse("Your report is available here - %s."
                        "\nThank you! In case of any questions contact yshutiko@redhat.com" % url)
