from django.shortcuts import render, redirect
from django.views.decorators.csrf import csrf_exempt
from .forms import DocumentForm
from .models import Document
from django.http import HttpResponse
from django.contrib import messages


# Create your views here.
@csrf_exempt
def upload_view(request):
    Document.objects.all().delete()  # clear legacy file

    message = "彙總表轉換器"

    # Handle file upload
    if request.method == 'POST':
        form = DocumentForm(request.POST, request.FILES)
        if form.is_valid():
            date = request.POST["roster_month"]  # search "roster_month" in the user POST request
            newdoc = Document(docfile=request.FILES['docfile'])  # assign variable "newdoc" to the uploaded Excel file
            newdoc.save()

            # Redirect to the document list after POST
            return excel_export(request, date)
            # function ENDS

        else:
            message = 'The form is not valid. Fix the following error:'

    else:
        form = DocumentForm()  # An empty, unbound form

    # Render list page with the documents and the form
    context = {'form': form, 'message': message}
    return render(request, 'list.html', context)


def excel_export(request, date):
    import pandas as pd
    import re
    from datetime import datetime
    import xlwt

    def get_timing(xlsx):
        final_list = []  # list storing all data like .dat file

        for i in range(0, len(xlsx["姓名"])):  # loop thru each employee, ROW

            for j in range(3, 34):  # Loop thru dates in a month, COLUMN
                temp_list_start = []  # sublist, initialize at each loop, START
                temp_list_end = []  # sublist, initialize at each loop, END

                try:
                    col_date = xlsx.columns.to_list()[j]  # set column name variable

                    check_null = xlsx[col_date].isnull()  # null check


                    if check_null[i] == False:
                        # start work
                        temp_list_start.append(xlsx["姓名"][i])  # Name
                        start_time = re.search("^[0-9].*", xlsx.iloc[i, j]).group()
                        temp_list_start.append(
                            f"{col_date}" + f"/{date[5:].lstrip('0')}/{date[0:4]}" + " " + f"{start_time}")  # Date + start time
                        final_list.append(temp_list_start)  # append to the final list, START

                        # end work
                        end_time = re.search(".+\n\Z", xlsx.iloc[i, j]).group()
                        ##### check if end time exists
                        if end_time.strip() != start_time:  # end_time exist
                            temp_list_end.append(xlsx["姓名"][i])  # Name

                            temp_list_end.append(
                                f"{col_date}" + f"/{date[5:].lstrip('0')}/{date[0:4]}" + " " + f"{end_time.strip()}")  # Date + end time
                            final_list.append(temp_list_end)  # append to the final list, END
                        else:
                            continue

                    else:
                        continue

                except:
                    break
        return final_list

    raw = str(Document.objects.all())
    raw = raw[22:-3]
    dir = f"/users/kelvin/desktop/pycharm/gols_duty_time/src/media/{raw}"

    df_new = pd.read_excel(dir, sheet_name="刷卡記錄",
                           skiprows=lambda x: x in [0, 1, 3])
    final_list = get_timing(df_new)  # call function `get_timing`

    # xlwt operation
    now = datetime.now()
    row = 0
    col = 0

    response = HttpResponse(content_type='application/ms-excel')
    response['Content-Disposition'] = f'attachment; filename="{int(now.timestamp())} .xls"'

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet("sheet2")

    for name, time in final_list:
        ws.write(row, col, name)
        ws.write(row, col + 1, time)
        row += 1

    wb.save(response)

    return response
