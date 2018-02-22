import PyPDF2
import csv
from pathlib import Path
import io
import pandas
import numpy
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
def Cpk(usl, lsl, avg, sigma , cf, sigma_cf):
    cpu = (usl - avg - (cf*sigma)) / (sigma_cf*sigma)
    cpl = (avg - lsl - (cf*sigma)) / (sigma_cf*sigma)
    cpk = numpy.min([cpu, cpl])
    return cpl,cpu,cpk
def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = io.BytesIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages,
                                  password=password,
                                  caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)

    text = retstr.getvalue()

    fp.close()
    device.close()
    retstr.close()
    return text
def filename_extraction(inp_filename):
    raw = inp_filename.split('_')
    dev = raw[1]
    volt = raw[2]
    temp = raw[3]
    condition = raw[4]+raw[5]+raw[6]+raw[7]
    return dev,volt,temp,condition

############################### User inputs ###############################################
path_of_files = r'C:\Users\vind\OneDrive - Cypress Semiconductor\documents\python_codes\EYE_DIAG_ANALYZER\pdf_ccg3pa2_tt'
pathlist = Path(path_of_files).glob('**/*.pdf')
output_filename = 'out'
automated_data_collection = 'yes' #'no'
cpl_matrix = []
cpu_matrix = []
cpk_matrix = []
################################# Program  Begins #########################################
if automated_data_collection == 'no':
    with open(output_filename +'raw'+ '.csv', 'a', newline='') as csvfile:
        mywriter1 = csv.DictWriter(csvfile, dialect='excel',
                                   fieldnames=['rise_time_average', 'rise_time_minimum', 'rise_time_maximum',
                                               'fall_time_average', 'fall_time_minimum', 'fall_time_maximum',
                                               'bit_rate_average', 'bit_rate_minimum', 'bit_rate_maximum',
                                               'voltage_swing_average', 'voltage_swing_minimum', 'voltage_swing_maximum', 'filename'])
        mywriter1.writeheader()
        for files in pathlist:
        ###################### extracting only measurement page of the pdf file ##########################################
            print(files.name)
            pdfFileObj = open(files,'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            pdfWriter = PyPDF2.PdfFileWriter()
            pdfReader.getNumPages()
            pageNum = 3
            pageObj = pdfReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)
            pdfOutput = open('temp.pdf', 'wb')
            pdfWriter.write(pdfOutput)
            pdfOutput.close()
        ######################### pdf to text conversion ################################
            x= convert_pdf_to_txt('temp.pdf')
            text_extracted = x.split()
            counter_list = list(enumerate(text_extracted, 1))
            rise_time_average = (counter_list[91])[1]
            fall_time_average =  (counter_list[93])[1]
            bit_rate_average = (counter_list[97])[1]
            rise_time_minimum = (counter_list[145])[1]
            fall_time_minimum =  (counter_list[147])[1]
            bit_rate_minimum = (counter_list[151])[1]
            rise_time_maximum = (counter_list[156])[1]
            fall_time_maximum =  (counter_list[158])[1]
            bit_rate_maximum = (counter_list[162])[1]
            voltage_swing_average = (counter_list[131])[1]
            voltage_swing_minimum = (counter_list[170])[1]
            voltage_swing_maximum = (counter_list[174])[1]
            data_raw = [float(rise_time_average), float(rise_time_minimum), float(rise_time_maximum), float(fall_time_average),
                        float(fall_time_minimum), float(fall_time_maximum), float(bit_rate_average), float(bit_rate_minimum),
                        float(bit_rate_maximum), float(voltage_swing_average), float(voltage_swing_minimum),
                        float(voltage_swing_maximum), files.name]
            print(data_raw)
            mywriter2 = csv.writer(csvfile, delimiter=',', dialect = 'excel')
            mywriter2.writerow(data_raw)
    ################## Analysis begins ##########################################
    pandas.set_option('display.expand_frame_repr', False)
    data = pandas.DataFrame.from_csv(output_filename + 'raw' +'.csv',index_col=None)
    data_grouped = data.agg([numpy.min, numpy.mean, numpy.max, numpy.std])
    print(data_grouped)
    writer = pandas.ExcelWriter(output_filename + '.xlsx')
    data_grouped.to_excel(writer, 'Sheet1')
    writer.save()
if automated_data_collection == 'yes':
    with open(output_filename +'raw'+ '.csv', 'a', newline='') as csvfile:
        mywriter1 = csv.DictWriter(csvfile, dialect='excel',
                                   fieldnames=['rise_time_average', 'rise_time_minimum', 'rise_time_maximum',
                                               'fall_time_average', 'fall_time_minimum', 'fall_time_maximum',
                                               'bit_rate_average', 'bit_rate_minimum', 'bit_rate_maximum',
                                               'voltage_swing_average', 'voltage_swing_minimum', 'voltage_swing_maximum', 'Device','Voltage','Temperature','Condition'])
        mywriter1.writeheader()
        for files in pathlist:
        ###################### extracting only measurement page of the pdf file ##########################################
            print(files.name)
            dev_no,v,t,cond = filename_extraction(files.name)
            pdfFileObj = open(files,'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
            pdfWriter = PyPDF2.PdfFileWriter()
            pdfReader.getNumPages()
            pageNum = 3
            pageObj = pdfReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)
            pdfOutput = open('temp.pdf', 'wb')
            pdfWriter.write(pdfOutput)
            pdfOutput.close()
        ######################### pdf to text conversion ################################
            x= convert_pdf_to_txt('temp.pdf')
            text_extracted = x.split()
            counter_list = list(enumerate(text_extracted, 1))
            rise_time_average = (counter_list[91])[1]
            fall_time_average =  (counter_list[93])[1]
            bit_rate_average = (counter_list[97])[1]
            rise_time_minimum = (counter_list[145])[1]
            fall_time_minimum =  (counter_list[147])[1]
            bit_rate_minimum = (counter_list[151])[1]
            rise_time_maximum = (counter_list[156])[1]
            fall_time_maximum =  (counter_list[158])[1]
            bit_rate_maximum = (counter_list[162])[1]
            voltage_swing_average = (counter_list[131])[1]
            voltage_swing_minimum = (counter_list[170])[1]
            voltage_swing_maximum = (counter_list[174])[1]
            data_raw = [float(rise_time_average), float(rise_time_minimum), float(rise_time_maximum), float(fall_time_average),
                        float(fall_time_minimum), float(fall_time_maximum), float(bit_rate_average), float(bit_rate_minimum),
                        float(bit_rate_maximum), float(voltage_swing_average), float(voltage_swing_minimum),
                        float(voltage_swing_maximum), dev_no, v,t,cond]
            print(data_raw)
            mywriter2 = csv.writer(csvfile, delimiter=',', dialect = 'excel')
            mywriter2.writerow(data_raw)
    ################## Analysis begins ##########################################
    pandas.set_option('display.expand_frame_repr', False)
    data = pandas.DataFrame.from_csv(output_filename + 'raw' +'.csv', index_col=None)
    data1 = data.groupby(['Voltage','Temperature','Condition'])
    data_grouped = data1.agg([numpy.min, numpy.mean, numpy.max, numpy.std])
    data_grouped2 = data1.agg([numpy.mean, numpy.std])
    # print(data_grouped2)
    for row in range(data_grouped2.shape[0]):
         slice = data_grouped2.iloc[row]
         # print(slice)
         for column in range(0,data_grouped2.shape[1],2):
             if column < 12:
                 mean1 = slice[column]
                 std1 = slice[column+1]
                 cpl,cpu,cpk = Cpk(800,300,mean1,std1,0,3)
                 cpk_matrix.append([cpl,cpu,cpk])
             if column >=12 and column <18:
                 mean1 = slice[column]
                 std1 = slice[column+1]
                 cpl, cpu, cpk = Cpk(320, 280, mean1, std1, 0, 3)
                 cpk_matrix.append([cpl,cpu,cpk])
             if column >= 18:
                 mean1 = slice[column]
                 std1 = slice[column+1]
                 cpl, cpu, cpk = Cpk(1.2, 1.05, mean1, std1, 0, 3)
                 cpk_matrix.append([cpl,cpu,cpk])
    cpk_df1 = pandas.DataFrame(cpk_matrix,columns = list(['Cpl','Cpu','Cpk']))
    cpk_df = numpy.array(cpk_df1)
    cpk_df_resize = numpy.resize(cpk_df,(data_grouped2.shape[0],36))
    # print(cpk_df_resize)
    sub_header = []
    field_hard = ['rise_time_average', 'rise_time_minimum', 'rise_time_maximum',
                    'fall_time_average', 'fall_time_minimum', 'fall_time_maximum',
                    'bit_rate_average', 'bit_rate_minimum', 'bit_rate_maximum',
                    'voltage_swing_average', 'voltage_swing_minimum', 'voltage_swing_maximum']
    for first in field_hard:
        for second in ['Cpl','Cpu','Cpk']:
            str_concat = second + '_' + first
            sub_header.append(str_concat)
    cpk_df2 = pandas.DataFrame(cpk_df_resize,columns= list(sub_header))
    # print(cpk_df2)
    data_grouped3 = pandas.concat([data_grouped2.reset_index(),cpk_df2],axis=1)
    data_grouped4 = pandas.concat([data_grouped.reset_index(),cpk_df2],axis=1)
    # print(data_grouped3)
    writer = pandas.ExcelWriter(output_filename + '.xlsx')
    data_grouped.to_excel(writer, 'Sheet1')
    data_grouped2.to_excel(writer, 'Sheet2')
    data_grouped3.to_excel(writer, 'Sheet3')
    data_grouped4.to_excel(writer, 'Sheet4')
    writer.save()
    print(data_grouped)
    print('Execution complete')
