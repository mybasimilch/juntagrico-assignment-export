from django.shortcuts import render
from django.contrib.auth.decorators import permission_required
from juntagrico.dao.subscriptiondao import SubscriptionDao

from django.utils.translation import gettext as _
from django.http import HttpResponse
from juntagrico.views import get_menu_dict
from io import BytesIO
from xlsxwriter import Workbook
from juntagrico.config import Config
from django.db.models import Sum, Case, When
from django.utils import timezone

from datetime import date,datetime,timedelta
from django.utils.dateparse import parse_date


def annotate_members_with_assignemnt_count(members, from_date, to_date):
    now = timezone.now()

    to_date = min(now, to_date)
    return members.annotate(assignment_count=Sum(
        Case(When(assignment__job__time__gte=from_date, assignment__job__time__lt=to_date, then='assignment__amount')))).annotate(
        core_assignment_count=Sum(Case(
            When(assignment__job__time__gte=from_date, assignment__job__time__lt=to_date, assignment__core_cache=True,
                 then='assignment__amount'))))


def subscriptions_with_assignments_during_timespan(subscriptions, from_date, to_date):
    subscriptions_list = []
    for subscription in subscriptions:
        assignments = 0
        core_assignments = 0
        members = annotate_members_with_assignemnt_count(
            subscription.members.all(), from_date, to_date)
        for member in members:
            assignments += member.assignment_count \
                if member.assignment_count is not None else 0
            core_assignments += member.core_assignment_count \
                if member.core_assignment_count is not None else 0
        subscriptions_list.append({
            'subscription': subscription,
            'assignments': assignments,
            'core_assignments': core_assignments
        })
    return subscriptions_list


@permission_required("juntagrico.is_operations_group")
def export_assignments(request):
    renderdict = get_menu_dict(request)
    renderdict.update({
        'menu': {'dm': 'active'},
        'change_date_disabled': True,
    })
    if request.method == 'POST':
        parsedFrom = parse_date(request.POST['fromDate'])
        parsedTo = parse_date(request.POST['toDate'])
        if not parsedFrom or not parsedTo:
            return render(request, "ae/export_overview.html", renderdict)
        fromDate = datetime.combine(parsedFrom, datetime.min.time(),tzinfo=timezone.utc)
        toDate = datetime.combine(parsedTo+timedelta(days=1), datetime.min.time(),tzinfo=timezone.utc)
        response = HttpResponse(
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = 'attachment; filename=Report.xlsx'
        output = BytesIO()
        workbook = Workbook(output)
        worksheet_s = workbook.add_worksheet(
            Config.vocabulary('subscription_pl'))

        worksheet_s.write_string(0, 0, str(_('Übersicht')))
        worksheet_s.write_string(0, 1, str(_('HauptbezieherIn')))
        worksheet_s.write_string(0, 2, str(_('HauptbezieherInEmail')))
        worksheet_s.write_string(0, 3, str(_('HauptbezieherInTelefon')))
        worksheet_s.write_string(0, 4, str(_('HauptbezieherInMobile')))
        worksheet_s.write_string(0, 5, str(_('Weitere BezieherInnen')))
        worksheet_s.write_string(0, 6, str(_('Depot')))
        worksheet_s.write_string(0, 7, str(Config.vocabulary('assignment_pl')))
        worksheet_s.write_string(0, 8, str(_('{} soll/Jahr'.format(Config.vocabulary('assignment_pl')))))
        

        subs = subscriptions_with_assignments_during_timespan(
            SubscriptionDao.all_active_subscritions(), fromDate, toDate)

        row = 1
        for sub in subs:
            worksheet_s.write_string(row, 0, sub['subscription'].overview)
            worksheet_s.write_string(
                row, 1, sub['subscription'].primary_member.get_name())
            worksheet_s.write_string(
                row, 2, sub['subscription'].primary_member.email)
            worksheet_s.write_string(
                row, 3, sub['subscription'].primary_member.phone or '')
            worksheet_s.write_string(
                row, 4, sub['subscription'].primary_member.mobile_phone or '')
            worksheet_s.write_string(
                row, 5, sub['subscription'].other_recipients_names())
            worksheet_s.write_string(row, 6, sub['subscription'].depot.name)
            worksheet_s.write(row, 7, sub.get('assignments'))
            worksheet_s.write(row, 8, sub['subscription'].required_assignments)
            row += 1

        workbook.close()
        xlsx_data = output.getvalue()
        response.write(xlsx_data)
        return response
    return render(request, "ae/export_overview.html", renderdict)
