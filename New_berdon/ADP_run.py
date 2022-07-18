from adp_py import *


path = r"C:\BerdonRPA\AuditSOC1\SOCFiles"
out=r"C:\BerdonRPA\AuditSOC1\ML\ADP_process"
Index=(out+'\\Index.pdf')
ind_txt=(out+'\\Index.txt')
processed=(out+'\\processed')
non_processed=(out+'\\non_processed')
report=(r'C:\BerdonRPA\AuditSOC1\ML\Report')
output_path=r'C:\BerdonRPA\AuditSOC1\ML\Extracted\ADP'
df_report= pd.DataFrame(columns=['Si_no','File_Name','Start_time','End_time','Duration', 'Result'])
try:
    shutil.rmtree(out)
except:
    pass
os.makedirs(out, exist_ok=False)
os.makedirs(processed, exist_ok=False)
os.makedirs(non_processed, exist_ok=False)
si=1
for filename in os.listdir(path):
    start_time = time.time()
    st= datetime.datetime.now()
    if filename.endswith(".pdf"):
        print(filename)
        pdf = pikepdf.open(os.path.join(path, filename))
        pdf.save(os.path.join(out, filename))#sequrity removed pdf saved
        pdf.close()
        pdf = PdfFileReader(os.path.join(out, filename), "rb")#read pdf
        pdf_writer = PdfFileWriter()
        for page in range(1, 2):
            pdf_writer.addPage(pdf.getPage(page))
        with open(Index, 'wb') as ind:#open new indexpdf file and write index page
            pdf_writer.write(ind)
        PDF_Parse = parser.from_file(Index)#read indexpdf file
        value=(PDF_Parse ['content'])
        with open(ind_txt, 'w') as f:
            f.write((value))#converted index pdf into txt
        res_dict, output_fname=index(ind_txt,filename,output_path,out,pdf)
        pdf2jpg.convert_pdf2jpg(output_fname, out, dpi=300, pages="ALL")
        objects,objects_table=detect(output_fname)
        valueDf=score(out,filename,objects)
        df_tabel=tabel(output_path,filename,objects_table)
        res_dict =resdict(valueDf,res_dict)
        res_dict= pd.DataFrame.from_dict(res_dict, orient="index").reset_index()
        writer = pd.ExcelWriter(os.path.join(output_path, filename[:-4]+'.xlsx'), engine='xlsxwriter')
        res_dict.to_excel(writer, sheet_name='Value')
        df_tabel.to_excel(writer, sheet_name='Table')
        writer.save()
        try:
            shutil.move(os.path.join(path, filename),processed)
            tt=((time.time() - start_time)/60)
            tt=round(tt,2)
        except:
            pass
        df_report = df_report.append({'Si_no':si,'File_Name': filename,'Start_time': (st.strftime('%H:%M:%S')),'End_time':(datetime.datetime.now().strftime('%H:%M:%S')),'Duration':tt, 'Result': "Processed"}, ignore_index=True)
        si=si+1 
    else:
        df_report = df_report.append({'Si_no':si,'File_Name': filename,'Start_time':(st.strftime('%H:%M:%S')),'End_time':(datetime.datetime.now().strftime('%H:%M:%S')),'Duration':'0', 'Result': "UN_Processed"}, ignore_index=True)
        si=si+1
        try:
            shutil.move(os.path.join(path, filename),non_processed)
        except:
            pass
df_report.to_excel(os.path.join(report,'ADP_Report_'+str(datetime.datetime.now().strftime('%m_%d_%Y_%H_%M_%S') ) +'.xlsx'))
print ("Report Generated")