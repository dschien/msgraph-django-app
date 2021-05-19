import openpyxl
from django.shortcuts import render
from django.http import HttpResponse, HttpResponseRedirect
from django.urls import reverse
from datetime import datetime, timedelta
from dateutil import tz, parser
from tutorial.auth_helper import get_sign_in_flow, get_token_from_code, store_user, remove_user_and_token, get_token
from tutorial.graph_helper import *
import dateutil.parser


# <HomeViewSnippet>
def home(request):
    context = initialize_context(request)

    return render(request, 'tutorial/home.html', context)


# </HomeViewSnippet>

# <InitializeContextSnippet>
def initialize_context(request):
    context = {}

    # Check for any errors in the session
    error = request.session.pop('flash_error', None)

    if error != None:
        context['errors'] = []
        context['errors'].append(error)

    # Check for user in the session
    context['user'] = request.session.get('user', {'is_authenticated': False})
    return context


# </InitializeContextSnippet>

# <SignInViewSnippet>
def sign_in(request):
    # Get the sign-in flow
    flow = get_sign_in_flow()
    # Save the expected flow so we can use it in the callback
    try:
        request.session['auth_flow'] = flow
    except Exception as e:
        print(e)
    # Redirect to the Azure sign-in page
    return HttpResponseRedirect(flow['auth_uri'])


# </SignInViewSnippet>

# <SignOutViewSnippet>
def sign_out(request):
    # Clear out the user and token
    remove_user_and_token(request)

    return HttpResponseRedirect(reverse('home'))


# </SignOutViewSnippet>

# <CallbackViewSnippet>
def callback(request):
    # Make the token request
    result = get_token_from_code(request)

    # Get the user's profile
    # user = get_user(result['code'])
    user = get_user(result['access_token'])

    # Store user
    store_user(request, user)
    return HttpResponseRedirect(reverse('home'))


# </CallbackViewSnippet>

# <CalendarViewSnippet>
def calendar(request):
    context = initialize_context(request)
    user = context['user']

    # Load the user's time zone
    # Microsoft Graph can return the user's time zone as either
    # a Windows time zone name or an IANA time zone identifier
    # Python datetime requires IANA, so convert Windows to IANA
    time_zone = get_iana_from_windows(user['timeZone'])
    tz_info = tz.gettz(time_zone)

    # Get midnight today in user's time zone
    today = datetime.now(tz_info).replace(
        hour=0,
        minute=0,
        second=0,
        microsecond=0)

    # Based on today, get the start of the week (Sunday)
    if (today.weekday() != 6):
        start = today - timedelta(days=today.isoweekday())
    else:
        start = today

    end = start + timedelta(days=7)

    token = get_token(request)

    events = get_calendar_events(
        token,
        start.isoformat(timespec='seconds'),
        end.isoformat(timespec='seconds'),
        user['timeZone'])

    if events:
        # Convert the ISO 8601 date times to a datetime object
        # This allows the Django template to format the value nicely
        for event in events['value']:
            event['start']['dateTime'] = parser.parse(event['start']['dateTime'])
            event['end']['dateTime'] = parser.parse(event['end']['dateTime'])

        context['events'] = events['value']

    return render(request, 'tutorial/calendar.html', context)


# </CalendarViewSnippet>

# <NewEventViewSnippet>
def newevent(request):
    context = initialize_context(request)
    user = context['user']

    if request.method == 'POST':
        # Validate the form values
        # Required values
        if (not request.POST['ev-subject']) or \
                (not request.POST['ev-start']) or \
                (not request.POST['ev-end']):
            context['errors'] = [
                {'message': 'Invalid values', 'debug': 'The subject, start, and end fields are required.'}
            ]
            return render(request, 'tutorial/newevent.html', context)

        attendees = None
        if request.POST['ev-attendees']:
            attendees = request.POST['ev-attendees'].split(';')
        body = request.POST['ev-body']

        # Create the event
        token = get_token(request)

        create_event(
            token,
            request.POST['ev-subject'],
            request.POST['ev-start'],
            request.POST['ev-end'],
            attendees,
            request.POST['ev-body'],
            user['timeZone'])

        # Redirect back to calendar view
        return HttpResponseRedirect(reverse('calendar'))
    else:
        # Render the form
        return render(request, 'tutorial/newevent.html', context)
    # print('hello')


# </NewEventViewSnippet>

def bulkevent(request):
    context = initialize_context(request)
    user = context['user']

    if request.method == 'POST':
        body = request.POST['ev-body']
        if not request.POST['ev-subject']:
            context['errors'] = [
                {'message': 'Invalid values', 'debug': 'The subject, start, and end fields are required.'}
            ]
            return render(request, 'tutorial/bulkevent.html', context)

        excel_file = request.FILES["excel_file"]
        # you may put validations here to check extension or file size
        try:
            data = read_excel(excel_file)
        except Exception as e:
            context['errors'] = [
                {'message': 'Excel parsing failed', 'debug': 'Check the format of your file.'}
            ]
            return render(request, 'tutorial/bulkevent.html', context)

        results = []
        for row in data:
            start_date = row[1]
            start_time = row[2]
            group = row[3]
            attendees = row[4:]
            # '2021-05-08T11:56'

            start_time = datetime.combine(dateutil.parser.parse(start_date).date(),
                                          dateutil.parser.parse(start_time).time()
                                          )

            end_time = start_time + timedelta(minutes=int(request.POST['ev-duration']))
            # Create the event
            token = get_token(request)

            res = create_event(
                token,
                request.POST['ev-subject'] + " " + group,
                start_time.isoformat(),
                end_time.isoformat(),
                attendees,
                request.POST['ev-body'],
                user['timeZone'])
            results.append({'result':res,'group': group})

        # Redirect back to calendar view

        context['messages'] = [
            {'message': f'Group {res["group"]}', 'detail': res["result"].status_code} for res in results
        ]
        return render(request, 'tutorial/bulkevent.html', context)
        # return HttpResponseRedirect(reverse('calendar'))
    else:
        # Render the form
        return render(request, 'tutorial/bulkevent.html', context)
    # print('hello')


def read_excel(excel_file):
    wb = openpyxl.load_workbook(excel_file)
    # getting a particular sheet by name out of many sheets
    worksheet = wb["schedule"]
    # print(worksheet)
    excel_data = list()
    # iterating over the rows and
    # getting value from each cell in row
    for row in worksheet.iter_rows():
        row_data = list()
        for cell in row:
            row_data.append(str(cell.value))
        excel_data.append(row_data)
    return excel_data
