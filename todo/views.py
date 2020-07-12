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

# PDF to Invoice
import stat
from pandas import DataFrame
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import BytesIO

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
    if 'upload_pdf' in request.POST:
        try:
            upload_file = request.FILES['document']
            uploaded_filename = upload_file.name
            print(uploaded_filename)
            print(upload_file.size)
            file_Extension = uploaded_filename.split(".")[-1]
            if file_Extension.lower() == "pdf":
                fspdf = FileSystemStorage(location='media/pdfs')
                fspdf.delete(upload_file.name)
                name = fspdf.save(upload_file.name, upload_file)
                context['url'] = fspdf.url('/pdfs/' + name)
            else:
                fs = FileSystemStorage()
                fs.delete(upload_file.name)
                name = fs.save(upload_file.name, upload_file)
                context['url'] = fs.url(name)
        except MultiValueDictKeyError:
            context['url'] = 'No file selected'


    elif 'delete_all_pdfs' in request.POST:
        fspdf_pdf = FileSystemStorage()
        pdf_path = fspdf_pdf.path('pdfs/')
        pdf_list = os.listdir(pdf_path)
        for pdf_file in pdf_list:
            fspdf_pdf.delete(pdf_path + "/" + pdf_file)
        context['url'] = "All PDFs Deleted"
        # return redirect('upload')

    return render(request, 'todo/upload.html', context)


@login_required
def reportgenerator(request):
    if request.method == 'GET':
        fspdf = FileSystemStorage()
        Saveto = fspdf.path('xlsx/')
        xlsx_list = os.listdir(Saveto)
        return render(request, 'todo/reportgenerator.html', {'xlsxlist': xlsx_list})


    elif 'delete_all' in request.POST:

        fspdf = FileSystemStorage()
        xlsx_path = fspdf.path('xlsx/')
        xlsx_list = os.listdir(xlsx_path)
        for xlsxs in xlsx_list:
            os.chmod(xlsx_path + "/" + xlsxs, stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH)
            fspdf.delete(xlsx_path + "/" + xlsxs)

        return redirect('reportgenerator')

    elif 'invoice' in request.POST:
        #try:
        def cn(n):
            try:
                return float(n)
            except:
                return 0
        fs = FileSystemStorage()
        def save_to_excel(writer, tableall):
            writer

            choice = 0
            if (tableall[0][0].isdigit()):
                choice = 0
            else:
                choice = 1

            if (choice == 0):
                t = tableall[0:11]
                df0 = DataFrame({'SHIP TO': [t[2][0:38], t[3][0:38], t[4][0:38], t[5][0:38], t[6][0:38]],
                                 'REMIT TO': [t[2][38:77], t[3][38:77], t[4][38:77], t[5][38:77], t[6][38:77]],
                                 'BILL TO': [t[2][77:105], t[3][77:105], t[4][77:105], t[5][77:105], t[6][77:105]],
                                 'INVOICE NO': [t[2][112:121], 'DATE', t[4][112:121], '', '']})

                df0.to_excel(writer, 'page' + t[0][0], startrow=0, startcol=0, index=False)

                df00 = DataFrame({t[7][0:38]: [], t[7][112:130]: []})
                df00.to_excel(writer, 'page' + t[0][0], startrow=7, startcol=0, index=False)

                df000 = DataFrame({'TAK': [t[9][0:6]],
                                   'STOP': [t[9][6:15]],
                                   'P.O. NUMBER': [t[9][15:29]],
                                   'SALESPERSON': [t[9][29:60]],
                                   'PRF#': [t[9][60:65]],
                                   'CUST. PHONE NO.': [t[9][65:77]],
                                   'DUNS NO.': t[9][77:91],
                                   'TERMS': t[9][91:130],
                                   })
                df000.to_excel(writer, 'page' + t[0][0], startrow=9, startcol=0, index=False)

                tableall = tableall[11:100]
                item = []
                UPC = []
                Quantity = []
                Description = []
                Packing = []
                Sugg_retail = []
                key = []
                GP = []
                R_E = []
                Unit_cost = []
                Cost_extension = []
                Tax = []
                for i in range(0, len(tableall)):
                    temp = tableall[i]
                    n = 0
                    item.append(temp[0:9])
                    UPC.append(temp[9:20])
                    Quantity.append(temp[21:29])
                    # print(Quantity)
                    Description.append(temp[30:63])
                    # print(Description)
                    Packing.append(temp[63:73])
                    # print(Packing)
                    Sugg_retail.append(temp[73:82])
                    # print(Sugg_retail)
                    key.append(temp[82:84])
                    # print(key)
                    GP.append(temp[86:91])
                    # print(GP)
                    R_E.append(temp[91:104])
                    # print(R_E)
                    Unit_cost.append(cn(temp[105:113]))
                    # print(Unit_cost)
                    Cost_extension.append(cn(temp[115:125]))
                    # print(Cost_extension)
                    Tax.append(temp[127:130])
                    # print(Tax)

                df = DataFrame(
                    {'ITEM': item, 'UPC': UPC, 'QUANTITY': Quantity, 'DESCRIPTION': Description, 'PACKING': Packing,
                     'SUGG RETAIL': Sugg_retail, 'KEY': key, 'GP%': GP, 'RETAIL EXTENSION': R_E,
                     'UNIT COST': Unit_cost, 'COST EXTENSION': Cost_extension, 'TAX': Tax})

                df.to_excel(writer, 'page' + t[0][0], startrow=11, startcol=0, index=False)

                worksheet = writer.sheets['page' + t[0][0]]  # pull worksheet object
                for idx, col in enumerate(df):  # loop through all columns
                    series = df[col]
                    max_len = max(( series.astype(str).map(len).max(),  # len of largest item
                        len(str(series.name))  # len of column name/header
                    )) + 1  # adding a little extra space
                    worksheet.set_column(idx, idx, max_len)  # set column width

            if (choice == 1):
                ttt = tableall
                check = 0
                tttt = []
                for i in range(0, len(ttt)):
                    for j in range(0, len(ttt[i])):
                        if (ttt[i][j].isdigit()):
                            check = check + 1
                            if (check >= 20):
                                tttt.append(ttt[i])
                                break
                CARTONS = tttt[0][1:9]
                CIG_TAX = tttt[0][9:18]
                NO_OF_LABEL = tttt[0][18:27]
                TAXABLE = tttt[0][27:40]
                NON_TAXABLE = tttt[0][40:51]
                GROCERY = tttt[0][51:63]
                NON_GROCERY = tttt[0][63:75]
                ALLOWANCES = tttt[0][75:100]
                INVOICE_TOTAL = tttt[0][115:130]

                LESS_APPLICABLE_DISCOUNT = tttt[1][115:130]
                PLUS_TOTAL_CHARGES = tttt[2][115:130]
                PLEASE_PAY_AMOUNT = tttt[3][115:130]

                TOTES = tttt[2][45:55]

                a = tttt[4][90:130]
                b = tttt[5][90:130]
                c = tttt[6][90:130]

                df2 = DataFrame({'CARTONS': [CARTONS, '', '', '', '', '', ''],
                                 'CIG TAX': [CIG_TAX, '', '', '', '', '', ''],
                                 'NO OF LABEL': [NO_OF_LABEL, '', '', '', '', '', ''],
                                 'TAXABLE': [TAXABLE, '', '', '', '', '', ''],
                                 'NON_TAXABLE': [NON_TAXABLE, '', '', '', 'TOTES IN', '', ''],
                                 'GROCERY': [GROCERY, '', '', '', TOTES, '', ''],
                                 'NON GROCERY': [NON_GROCERY, '', '', '', '', '', ''],
                                 'ALLOWANCES': [ALLOWANCES, '', '', '', '', '', ''],
                                 '': ['', '', '', '', '', '', ''],
                                 'INVOICE TOTAL': ['LESS APPLICABLE DISCOUNT', 'PLUS TOTAL CHARGES',
                                                   'PLEASE PAY THIS AMOUNT', '', a, b, c],
                                 INVOICE_TOTAL: [LESS_APPLICABLE_DISCOUNT, PLUS_TOTAL_CHARGES, PLEASE_PAY_AMOUNT,
                                                 '', '', '', '']})

                df2.to_excel(writer, sheet_name='total', startrow=0, startcol=0, index=False)

        def convert_pdf(pdf_path, codec='utf-8', password=''):

            rsrcmgr = PDFResourceManager()
            retstr = BytesIO()
            laparams = LAParams()
            device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
            fp = open(pdf_path, 'rb')
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            maxpages = 0
            caching = True
            pagenos = set()

            for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                          check_extractable=True):
                interpreter.process_page(page)

            text = retstr.getvalue().decode()
            fp.close()
            device.close()
            retstr.close()
            return text

        def text_to_table(text):
            pages = text.split('                                                                   000')
            #    tables_all=[]
            pagess = []
            for i in range(1, len(pages)):
                page_text = pages[i].split('\n \n \n \n \n')
                f = []
                check = 0
                for k in range(0, len(page_text) - 1):

                    for j in range(0, len(page_text[k])):
                        if (page_text[k][j].isdigit()):
                            check = check + 1
                            if (check >= 20):
                                f.append(k)
                                break

                    if (len(f) == 2):
                        #            total=page_text[f[1]]
                        pagess.append(page_text[f[1]])
                        del f[1]

                pagess.append(page_text[f[0]])
            return pagess

        def getPfromT(tableall):
            #    tableall=tableall[11:100]
            item = []
            UPC = []
            Quantity = []
            Description = []
            Packing = []
            Sugg_retail = []
            key = []
            GP = []
            R_E = []
            Unit_cost = []
            Cost_extension = []
            Tax = []
            Discount = []
            p = []
            n = 10
            for i in range(0, len(tableall)):
                temp = tableall[i]
                if (temp[9:20].isdigit()):
                    if ('CT ' in temp[21:63]):
                        # if(True):
                        item.append(temp[0:9])
                        UPC.append(temp[9:20])
                        Quantity.append(temp[21:29])
                        # print(Quantity)
                        Description.append(temp[30:63])
                        # print(Description)
                        Packing.append(temp[63:73])
                        if (('25' in temp[63:73]) | ('8/' in temp[63:73])):
                            p.append(8)
                        elif (('20' in temp[63:73])):
                            p.append(10)
                        elif (temp[63:73] == ' 10/8S    '):
                            p.append(10)
                        else:
                            p.append(0)

                        # print(Packing)
                        Sugg_retail.append(temp[73:82])
                        # print(Sugg_retail)
                        key.append(temp[82:84])
                        # print(key)
                        GP.append(temp[86:91])
                        # print(GP)
                        R_E.append(temp[91:104])
                        # print(R_E)
                        Unit_cost.append(cn(temp[105:113]))
                        # print(Unit_cost)
                        Cost_extension.append(cn(temp[115:125]))
                        # print(Cost_extension)
                        Tax.append(temp[127:130])
                        # print(Tax)
                        n = 0
                        Discount.append(0.00)

                elif (n == 0):
                    Discount[len(Discount) - 1] = Discount[len(Discount) - 1] + cn(temp[105:113])
                    n = 1
                elif (n == 1):
                    Discount[len(Discount) - 1] = Discount[len(Discount) - 1] + cn(temp[105:113])
                    n = 2

            df = DataFrame({'ITEM': item,
                            'UPC': UPC,
                            'QUANTITY': Quantity,
                            'DESCRIPTION': Description,
                            'PACKING': Packing,
                            'P': p,
                            'UNIT COST': Unit_cost,
                            'Unit Disc': Discount})
            return df

        def convertMultiple(pdfDir, txtDir):
            #    if pdfDir == "": pdfDir = os.getcwd() + "\\" #if no pdfDir passed in
            df = DataFrame({'ITEM': [],
                            'UPC': [],
                            'QUANTITY': [],
                            'DESCRIPTION': [],
                            'PACKING': [],
                            'P': [],
                            'Date': [],
                            'UNIT COST': [],
                            'Unit Disc': []
                            })
            #    print(df)
            for pdf in os.listdir(pdfDir):  # iterate through pdfs in pdf directory
                # print(pdf)
                fileExtension = pdf.split(".")[-1]
                fileName = pdf.split(".")[0]
                if fileExtension.lower() == "pdf":
                    pdfFilename = pdfDir + "/" + pdf
                    text = convert_pdf(pdfFilename)  # get string of text content of pdf
                    print("txtDirlas " + txtDir)
                    tables = text_to_table(text)
                    writer = pd.ExcelWriter(txtDir + "/" + fileName+'.xlsx',engine='xlsxwriter')
                    # writer = pd.ExcelWriter(fs.path('xlsx') + fileName+'.xlsx', engine='xlsxwriter')
                    try:
                        os.chmod(txtDir + "/" + fileName + '.xlsx', stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH)
                        # os.chmod(fs.path('table.xlsx'), stat.S_IWUSR | stat.S_IWGRP | stat.S_IWOTH)
                        # print('Files Created Las')
                    except:
                        print('Files Created')
                    tableall = ['']
                    for i in range(0, len(tables)):
                        table = tables[i].split('\n')
                        tableall = tableall + table[11:100]
                        date = table[4][112:121]
                        #                len(tableall)
                        save_to_excel(writer, table)
                    df_temp = getPfromT(tableall)
                    df_temp['Date'] = date;
                    df = df.append(df_temp, sort=False, ignore_index=True)
                    df = df[['ITEM', 'UPC', 'QUANTITY', 'DESCRIPTION', 'PACKING', 'P', 'Date', 'UNIT COST', 'Unit Disc']]
                    writer.save()
                    writer.close()
                    os.chmod(txtDir + "/" + fileName + '.xlsx', stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)
                    # os.chmod(fs.path('table.xlsx'), stat.S_IREAD | stat.S_IRGRP | stat.S_IROTH)  # to make it read only

            Margin = pd.read_excel(fs.path('table.xlsx'), index_col=0)
            m_temp = Margin.values.tolist()
            margin1 = m_temp[0][0]
            margin2 = m_temp[1][0]
            margin3 = m_temp[2][0]

            df['Net Cost'] = (df['UNIT COST']) - (df['Unit Disc'])
            df['Cost of Single'] = (df['Net Cost']) / (df['P'])
            df['Price of Single'] = margin1 * (df['Cost of Single']) + (df['Cost of Single'])
            df['Price of 2-pack'] = 2 * (margin2 * (df['Cost of Single']) + (df['Cost of Single']))
            df['Price of Carton'] = margin3 * (df['Net Cost']) + df['Net Cost']
            #    df['Net Cost']=(df['UNIT COST']) - (df['Unit Disc'])
            #    print(df)
            df = df.sort_values('Date', ascending=True)
            df = df.drop_duplicates(subset=['UPC', 'Net Cost'], keep='last')

            writer1 = pd.ExcelWriter(fs.path('table.xlsx'), engine='xlsxwriter')
            df.to_excel(writer1, sheet_name='all', startrow=0, startcol=0, index=False)

            workbook = writer1.book
            worksheet = writer1.sheets['all']
            # Add a number format for cells with money.
            money_fmt = workbook.add_format({'num_format': '$0.00', 'bold': True})
            worksheet.set_column('H:N', 13.57, money_fmt)
            worksheet.set_column('B:B', 11.3)
            worksheet.set_column('C:C', 10)
            worksheet.set_column('D:D', 36)
            worksheet.set_column('F:F', 2)
            # df.style.format({'Cost of Single':lambda x: money(x)})
            return df

        PDFFolder = fs.path('pdfs/') # os.getcwd() + '/'
        Saveto = fs.path('xlsx/') # os.getcwd() + '/xlsx/'
        df = convertMultiple(PDFFolder, Saveto)
        messages.add_message(request, messages.INFO, 'PDFs Converted to Excel Successfully!')
        return redirect('reportgenerator')
        # except ValueError:
            # return render(request, 'todo/reportgenerator.html', {'form': TodoForm(), 'error': 'Bad data passed in. Try again.'})

    elif 'comparison' in request.POST:
        try:
            # form = TodoForm(request.POST)
            # newtodo = form.save(commit=False)
            # newtodo.user = request.user
            # newtodo.save()
            # cwd = os.getcwd()
            #itemfilepath = settings.STATIC_ROOT, 'reports/4101_ItemFileReportRetailer_MAY31.xls'
            #itemfilepath = os.path.join(settings.BASE_DIR, '/reports/4101_ItemFileReportRetailer_MAY31.xls')
            # itemfilepath = '4101_ItemFileReportRetailer_MAY31.xls'
            # itemfilepath = cwd + '\\todo\\static\\reports\\4101_ItemFileReportRetailer_MAY31.xls'
            fs = FileSystemStorage()
            itemfilepath = fs.path('4101_ItemFileReportRetailer_MAY31.xls')
            df4101 = pd.read_excel(itemfilepath)

            itlpath = fs.path('ITL.xlsx')
            print("laspath" + itlpath)
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

        except ValueError:
             return render(request, 'todo/reportgenerator.html', {'form': TodoForm(), 'error': 'Bad data passed in. Try again.'})

