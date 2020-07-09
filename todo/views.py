import os

from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django.contrib.auth.models import User
from django.db import IntegrityError
from django.contrib.auth import login, logout, authenticate
from django.utils.datastructures import MultiValueDictKeyError

from todowoo import settings
from .forms import TodoForm
from .models import Todo
from django.utils import timezone
from django.contrib.auth.decorators import login_required
import pandas as pd
from django.contrib import messages
from django.templatetags.static import static
from django.core.files.storage import FileSystemStorage


def home(request):
    return render(request, 'todo/home.html')


def signupuser(request):
    if request.method == 'GET':
        return render(request, 'todo/signupuser.html', {'form': UserCreationForm()})
    else:
        if request.POST['password1'] == request.POST['password2']:
            try:
                user = User.objects.create_user(request.POST['username'], password=request.POST['password1'])
                user.save()
                login(request, user)
                return redirect('currenttodos')
            except IntegrityError:
                return render(request, 'todo/signupuser.html', {'form':UserCreationForm(), 'error':'That username has already been taken. Please choose a new username'})
        else:
            return render(request, 'todo/signupuser.html', {'form':UserCreationForm(), 'error':'Passwords did not match'})


def loginuser(request):
    if request.method == 'GET':
        return render(request, 'todo/loginuser.html', {'form':AuthenticationForm()})
    else:
        user = authenticate(request, username=request.POST['username'], password=request.POST['password'])
        if user is None:
            return render(request, 'todo/loginuser.html', {'form':AuthenticationForm(), 'error':'Username and password did not match'})
        else:
            login(request, user)
            return redirect('currenttodos')

@login_required
def logoutuser(request):
    if request.method == 'POST':
        logout(request)
        return redirect('home')

@login_required
def createtodo(request):
    if request.method == 'GET':
        return render(request, 'todo/createtodo.html', {'form':TodoForm()})
    else:
        try:
            form = TodoForm(request.POST)
            newtodo = form.save(commit=False)
            newtodo.user = request.user
            newtodo.save()
            return redirect('currenttodos')
        except ValueError:
            return render(request, 'todo/createtodo.html', {'form':TodoForm(), 'error':'Bad data passed in. Try again.'})

@login_required
def currenttodos(request):
    todos = Todo.objects.filter(user=request.user, datecompleted__isnull=True)
    return render(request, 'todo/currenttodos.html', {'todos':todos})

@login_required
def completedtodos(request):
    todos = Todo.objects.filter(user=request.user, datecompleted__isnull=False).order_by('-datecompleted')
    return render(request, 'todo/completedtodos.html', {'todos':todos})

@login_required
def viewtodo(request, todo_pk):
    todo = get_object_or_404(Todo, pk=todo_pk, user=request.user)
    if request.method == 'GET':
        form = TodoForm(instance=todo)
        return render(request, 'todo/viewtodo.html', {'todo':todo, 'form':form})
    else:
        try:
            form = TodoForm(request.POST, instance=todo)
            form.save()
            return redirect('currenttodos')
        except ValueError:
            return render(request, 'todo/viewtodo.html', {'todo':todo, 'form':form, 'error':'Bad info'})

@login_required
def completetodo(request, todo_pk):
    todo = get_object_or_404(Todo, pk=todo_pk, user=request.user)
    if request.method == 'POST':
        todo.datecompleted = timezone.now()
        todo.save()
        return redirect('currenttodos')

@login_required
def deletetodo(request, todo_pk):
    todo = get_object_or_404(Todo, pk=todo_pk, user=request.user)
    if request.method == 'POST':
        todo.delete()
        return redirect('currenttodos')


@login_required
def upload(request):
    todos = Todo.objects.filter(user=request.user)
    context = {}
    if request.method == 'POST':
        try:
            upload_file = request.FILES['document']
            print(upload_file.name)
            print(upload_file.size)
            fs = FileSystemStorage()
            name = fs.save(upload_file.name, upload_file)
            context['url'] = fs.url(name)
        except MultiValueDictKeyError:
            context['url'] = 'No'

    return render(request, 'todo/upload.html', context)


@login_required
def reportgenerator(request):
    if request.method == 'GET':
        return render(request, 'todo/reportgenerator.html', {'form': TodoForm()})
    else:
        #try:
        # form = TodoForm(request.POST)
        # newtodo = form.save(commit=False)
        # newtodo.user = request.user
        # newtodo.save()
        fs = FileSystemStorage()
        cwd = os.getcwd()
        #itemfilepath = settings.STATIC_ROOT, 'reports/4101_ItemFileReportRetailer_MAY31.xls'
        #itemfilepath = os.path.join(settings.BASE_DIR, '/reports/4101_ItemFileReportRetailer_MAY31.xls')
        # itemfilepath = '4101_ItemFileReportRetailer_MAY31.xls'
        # itemfilepath = cwd + '\\todo\\static\\reports\\4101_ItemFileReportRetailer_MAY31.xls'
        itemfilepath = fs.path('4101_ItemFileReportRetailer_MAY31.xls')
        df4101 = pd.read_excel(itemfilepath)

        itlpath = fs.path('ITL.xlsx')
        # itlpath = cwd + '\\todo\\static\\reports\\ITL.xlsx'
        # itlpath = os.path.join(settings.STATIC_URL, 'reports/ITL.xlsx')
        # itlpath = os.path.join(settings.MEDIA_ROOT, '/ITL.xlsx')
        # itlpath = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'static/reports/ITL.xlsx')
        # itlpath = 'ITL.xlsx'
        dfttl = pd.read_excel(itlpath)

        # coremarkpath = os.path.join(settings.STATIC_URL, 'reports/COREMARK.xlsx')
        # coremarkpath = cwd + '\\todo\\static\\reports\\COREMARK.xlsx'
        coremarkpath = fs.path('COREMARK.xlsx')
        dfcmt = pd.read_excel(coremarkpath, sheet_name='TOBACCO')

        # dfcmv=pd.read_excel('coremark.xlsx',sheet_name='VAPE')
        # dfcmc=pd.read_excel('coremark.xlsx',sheet_name='CIGAR')
        # dfcmo=pd.read_excel('coremark.xlsx',sheet_name='OTHER')
        dfcmt = dfcmt.append(pd.DataFrame({'Supplier': []}), sort=False)
        dfcmt['Supplier'] = 'COREMARK'
        df1 = dfcmt.iloc[1:len(dfcmt), [0, 1, 3, 27, 5, 7, 17, 9, 11]]
        df1 = df1.rename(index=str,
                         columns={"UNIT COST": "CC", "Unnamed: 17": "CR", "SINGLE": "SC", "Unnamed: 11": "SR"})

        dfttl = dfttl.append(pd.DataFrame({'Supplier': []}), sort=False)
        dfttl['Supplier'] = 'ITL'
        df2 = dfttl.iloc[1:len(dfttl), [0, 1, 2, 31, 3, 19, 21, 13, 15]]
        df2 = df2.rename(index=str,
                         columns={"CARTON": "CC", "Unnamed: 21": "CR", "SINGLE": "SC", "Unnamed: 15": "SR"})
        # dft=df.iloc[9:243,3:15]
        df0 = df1.append(df2)
        # df0['SUPPLIER CODE'] = pd.to_numeric(df0['SUPPLIER CODE'], errors='coerce')
        df0['SUPPLIER CODE'] = df0['SUPPLIER CODE'].astype(float).astype(int).astype(str)
        # df0['SUPPLIER CODE'] =df0['SUPPLIER CODE'].apply(str)
        df0['CC'] = df0['CC'].astype(float)
        df0['CR'] = df0['CR'].astype(float)
        df0['SC'] = df0['SC'].astype(float)
        df0['SR'] = df0['SR'].astype(float)
        # sizedf0=len(df0)-1
        df0 = df0.append(pd.DataFrame({'CCr': []}), sort=False)  # Carton Cost report
        df0 = df0.append(pd.DataFrame({'CRr': []}), sort=False)  # Carton Retail report
        df0 = df0.append(pd.DataFrame({'SCr': []}), sort=False)  # Single Cost report
        df0 = df0.append(pd.DataFrame({'SRr': []}), sort=False)  # Single Retail report
        df0 = df0.append(pd.DataFrame({'PN': []}), sort=False)  # CT product number
        df0 = df0.append(pd.DataFrame({'DS': []}), sort=False)  # Difference single
        df0 = df0.append(pd.DataFrame({'DC': []}), sort=False)  # Difference Carton
        df0 = df0.append(pd.DataFrame({'DM': []}), sort=False)  # PB margin
        # df0=df0.append(pd.DataFrame({'PN':[]}),sort=False) # CT product number

        for i, SupplierCode in enumerate(df0.iloc[:, 0]):

            dftemp = df4101[df4101['Unnamed: 11'] == SupplierCode]
            for j, v in enumerate(dftemp.iloc[:, 5]):
                if (v == 1):
                    df0.iloc[i, 13] = dftemp.iloc[j, 3]
                    df0.iloc[i, 11] = dftemp.iloc[j, 9]
                    df0.iloc[i, 12] = dftemp.iloc[j, 8]
                elif (v == 8):
                    df0.iloc[i, 9] = dftemp.iloc[j, 9]
                    df0.iloc[i, 10] = dftemp.iloc[j, 8]

                elif (v == 10):
                    df0.iloc[i, 9] = dftemp.iloc[j, 9]
                    df0.iloc[i, 10] = dftemp.iloc[j, 8]

                else:
                    print('No match')

        df0.iloc[:, 14] = df0['SC'] - df0['SCr']
        df0.iloc[:, 15] = df0['CC'] - df0['CCr']
        df0.iloc[:, 16] = (df0['CRr'] - df0['CCr']) / df0['CRr']
        df = pd.DataFrame([])
        df['SUPPLIER CODE'] = df0['SUPPLIER CODE']
        df['UPC'] = df0['UPC']
        df['CT PRODUCT #'] = df0['PN']
        df['DESCRIPTION'] = df0['DESCRIPTION']
        df['PACKING'] = df0['P']
        df['SUPPLIER'] = df0['Supplier']
        df['INVOICE COST'] = df0['CC']
        df['PRICE BOOK COST (4101)'] = df0['CCr']
        df['DIFFERENCE'] = df0['DC']
        df['PB MARGIN %'] = df0['DM']
        df = df.append(pd.DataFrame({'SUMMARY': []}), sort=False)
        df['INVOICE SELLING PRICE'] = df0['SC']
        df['PRICE BOOK RETAIL PRICE (4101)'] = df0['SCr']
        df['DIFFERENCE '] = df0['DS']
        df = df.append(pd.DataFrame({'NOTE': []}), sort=False)
        df.iloc[:, 6:16] = df.iloc[:, 6:16].round(2)

        # writer = pd.ExcelWriter(cwd +'\\todo\\static\\reports\\Comparison_Report.xlsx', engine='xlsxwriter')
        writer = pd.ExcelWriter(fs.path('Comparison_Report.xlsx'), engine='xlsxwriter')
        df.to_excel(writer, 'TOBACCO', startrow=1, startcol=0, index=False)
        workbook = writer.book
        worksheet = writer.sheets['TOBACCO']
        header_format = workbook.add_format({
            'align': 'center',
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#336699',
            'font_color': 'white',
            'border': 1})

        for col_num, value in enumerate(df.columns.values):
            if col_num < 6:
                worksheet.merge_range(0, col_num, 1, col_num, value, header_format)
        worksheet.merge_range(0, 6, 0, 10, value, header_format)
        worksheet.merge_range(0, 11, 0, 14, value, header_format)

        for col_num, value in enumerate(df.columns.values):
            if col_num > 5:
                worksheet.write(1, col_num, value, header_format)
        money_fmt = workbook.add_format({'align': 'center', 'num_format': '$0.00', 'bold': False})
        Quan = workbook.add_format({'align': 'center', 'bold': False})
        worksheet.set_column('A:A', 10, Quan)
        worksheet.set_column('B:B', 11.3, Quan)
        worksheet.set_column('C:C', 10, Quan)
        worksheet.set_column('D:D', 33)
        worksheet.set_column('E:E', 11, Quan)
        worksheet.set_column('F:F', 13)
        worksheet.set_column('G:G', 13, money_fmt)
        worksheet.set_column('H:H', 22.5, money_fmt)
        worksheet.set_column('I:I', 11, money_fmt)
        worksheet.set_column('J:J', 13, money_fmt)
        worksheet.set_column('K:K', 10, money_fmt)
        worksheet.set_column('L:L', 22, money_fmt)
        worksheet.set_column('M:M', 29, money_fmt)
        worksheet.set_column('N:N', 12, money_fmt)
        worksheet.set_column('O:O', 5.2, money_fmt)
        writer.save()
        messages.add_message(request, messages.INFO, 'Comparison Report Generated Successfully!')
        # messagebox.showinfo('Info', 'Comparision Report Generated Successfully!')
        # worksheet.set_column('O:AA', 9,money_fmt)
        # df['NOTE']=df0[]
        # input('\nPlease press enter')
        return redirect('reportgenerator')

        # except ValueError:
        #     return render(request, 'todo/reportgenerator.html',
        #                   {'form': TodoForm(), 'error': 'Bad data passed in. Try again.'})