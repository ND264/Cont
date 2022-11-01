#!/usr/bin/env python3
import pandas as pd
import time
from fpdf import FPDF
# Добавляем необходимые подклассы - MIME-типы
from email.mime.multipart import MIMEMultipart  # Многокомпонентный объект
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import smtplib


# Start !!!!!!!!!!!!!
def main():
    path_baza = "/Users/nd/Desktop/baza/baza.xlsm"
    global report
    report = []
    name_cols = ['Nr.', 'Denumirea', 'Contract', 'Taxa', 'Datorie_31.12.2021',
                 'Servicie_t1', 'Extrase_t1', 'Lista_t1', 'Scrisori_t1',
                 'Emisie_t1', 'Informatie_t1', 'Consulting_t1', 'PO_t1',
                 'Total_t1', 'Achitat_t1', 'Datorie_t1', 'Servicie_t2',
                 'Extrase_t2', 'Lista_t2', 'Scrisori_t2',
                 'Emisie_t2', 'Informatie_t2', 'Consulting_t2', 'PO_t2',
                 'Total_t2', 'Achitat_t2', 'Datorie_t2', 'Servicie_t3',
                 'Extrase_t3', 'Lista_t3', 'Scrisori_t3',
                 'Emisie_t3', 'Informatie_t3', 'Consulting_t3', 'PO_t3',
                 'Total_t3', 'Achitat_t3', 'Datorie_t3', 'Servicie_t4',
                 'Extrase_t4', 'Lista_t4', 'Scrisori_t4',
                 'Emisie_t4', 'Informatie_t4', 'Consulting_t4', 'PO_t4',
                 'Total_t4', 'Achitat_t4', 'Datorie_t4']

    global table, table_email
    table = pd.DataFrame(
        pd.read_excel(path_baza, sheet_name='2022', header=None))
    table_email = pd.DataFrame(
        pd.read_excel(path_baza, sheet_name='data', header=None))
    print(f"len={len(table)}, последнее АО: {table.loc[len(table) - 2][1]}")
    # table.drop(index=[0, 1, len(table) - 1], axis=0,
    #            inplace=True)  # удаляем строку
    # table.drop(columns=[0, 3], axis=1, inplace=True)  # удаляем столбец

    # print(f"len={len(table_email)}")
    print(table.to_string())  # Вывод всей базы
    # print(table_email.to_string())
    # em = "VIV"
    # em = "ACCENT ELECTRONIC"
    # em = "ZIMBRU-NORD"
    # em = "ZODIER"
    # em='AGENŢIA "CORCIMARU, PARTENERII ȘI ASOCIAŢII"'
    # em = "AGROSERV"

    print("""Формирование и отправка счетов
    1. Счёт одному АО за текущий квартал и/или Счёт до конца года
    2. Счёта всем АО за текущий квартал и/или Счёт до конца года
    3. Информация
    0. Выход""")
    choice = input("Что делаем: ")
    if choice == "1":
        cont_for_one()
    elif choice == "2":
        conts_for_all()
    elif choice == "3":
        infos()


def search_SA(emitent):
    sa_datas = []
    poz1 = None
    # Поиск АО
    for i in range(2, len(table)):
        if table.loc[i][1] == emitent:
            poz1 = i
            sa_datas.append(i)
            # print("poz1=", i)
            for j in range(1, len(table_email)):
                if table_email.loc[j][1] == emitent:
                    # poz2 = j
                    # print("poz2=", j)
                    sa_datas.append(table_email.loc[j][8])
                    break
            break
    if poz1 is None:
        return False
    else:
        return sa_datas

    # print("kkkkk====", table.loc[poz1][1], table.loc[poz1][49], "email:",
    #       table_email.loc[poz2][8])
    # print(table_email.to_string())
    # create_pdf(poz1, poz2, 3, True)
    # print(table.loc[table[1] == em])

    # bbb=table.loc[table[1]==em]
    # print(bbb[1], bbb[2], bbb[4])
    # for i in bbb.to_string():
    #     print(i)
    # wbFile = openpyxl.load_workbook(filename=path_baza, data_only=True)
    # row_max = wbFile["2022"].max_row
    # row_max2 = wbFile["data"].max_row
    # for i in report:
    #	print(i)


def create_pdf(p1, p2, trim, year=True, send_email=False):
    """
    Функция создающая pdf-документ и записывает файл
    :param send_email:
    :param p1: Номер строки где находится АО на листе "2022"
    :param p2: Адрес электронной почты
    :param trim: номер квартала 1-4 или 0 если счёт на год
    :param year: true/false если нужен 2-й счет до конца года
    :return:
    """
    period = [{'Servicii_de_registrator_Anul_2022': 49},
              {'Datorie': 5, 'Servicii_de_tinere_a_registrului_trim.1': 6,
               'Extras': 7, 'Lista': 8, 'Scrisori': 9, 'Emisie': 10,
               'Informatie': 11, 'Consalting': 12, 'Alte': 13,
               'Spre_achitare': 16},
              {'Datorie': 16, 'Servicii_de_tinere_a_registrului_trim.2': 17,
               'Extras': 18, 'Lista': 19, 'Scrisori': 20, 'Emisie': 21,
               'Informatie': 22, 'Consalting': 23, 'Alte': 24,
               'Spre_achitare': 27},
              {'Datorie': 27, 'Servicii_de_tinere_a_registrului_trim.3': 28,
               'Extras': 29, 'Lista': 30, 'Scrisori': 31, 'Emisie': 32,
               'Informatie': 33, 'Consalting': 34, 'Alte': 35,
               'Spre_achitare': 38},
              {'Datorie': 38, 'Servicii_de_tinere_a_registrului_trim.4': 39,
               'Extras': 40, 'Lista': 41, 'Scrisori': 42, 'Emisie': 43,
               'Informatie': 44, 'Consalting': 45, 'Alte': 46,
               'Spre_achitare': 49}]
    data = []
    year = False if trim == 4 else year
    xS = 100  # расположение печати по оси X
    yS = -3  # расположение печати по оси Y
    for k, v in period[int(trim)].items():
        if str(table.loc[p1][v]) != 'nan':
            # print(type(table.loc[p1][v]))
            if k == "Datorie" and table.loc[p1][v] < 0:
                k = "Avans"
            data.append([k, table.loc[p1][v]])
            yS += 1
            # print("key:", k, "value:", table.loc[p1][v])
    # print("DATA: ", data)

    date = time.strftime("%d.%m.%Y")
    txt1 = f"CONT DE PLATA din {date}"
    txt2 = u'Furnizor: S.A. "Grupa Financiară", IDNO: 1002600054323'
    txt3 = 'Adresa: MD2001, mun. Chişinău, str. A.Bernardazzi, 7, of. 7'
    txt4 = 'IBAN: MD17AG000000022515490120 la B.C. "MOLDOVA-AGROINDBANK" S.A. suc. Tighina'
    txt5 = 'Cod bancar: AGRNMD2X864'
    txt6 = u'Platitor: S.A. "' + table.loc[p1][1] + '"'
    txt_ps1 = u"""Если у Вашей компании есть возможность оплатить услуги за оставшийся период 2022 года, ниже представлен счёт по которому можно произвести платёж."""
    txt_ps2 = u"""ВНИМАНИЕ!!! Налоговая накладная на наши услуги отписывается посредством системы 'E-FACTURA' по длинному циклу. Просим своевременно подписывать её. Все вопросы и пожелания принимаются по телефону 022-27-23-13 (бухгалтерия)."""
    txt_ps3 = u"""IN ATENTIA CONTABILITATII!!! Factura fiscala pentru serviciile acordate SA 'Grupa Financiară' este perfectata prin intermediul 'E-FACTURA' - ciclu mare. Rugam semnarea acestei facturi in timp util. Ralatii la telefon 022-27-37-13 (contabilitatea)."""
    txt_ps4 = u"""Dacă doriți să achitați serviciile totale pentruanul 2022 mai jos va prezentăm contul anual."""
    txt_sign1 = 'Directoare Viorica BONDAREV'
    txt_sign2 = u'Contabila-şef Maia CIULCOVA'

    pdf = FPDF()
    pdf.add_page()
    pdf.add_font('DejaVu', '',
                 '/Users/nd/Desktop/baza/fonts/DejaVuSansCondensed.ttf',
                 uni=True)
    pdf.add_font('DejaVu', 'B',
                 '/Users/nd/Desktop/baza/fonts/DejaVuSans-Bold.ttf', uni=True)

    pdf.set_font("DejaVu", size=20, style="B")
    pdf.cell(200, 10, txt1, ln=1, align="C")
    pdf.set_font("DejaVu", size=13)
    pdf.cell(200, 6, txt2, ln=1)
    pdf.cell(200, 6, txt3, ln=1)
    pdf.cell(200, 6, txt4, ln=1)
    pdf.cell(200, 6, txt5, ln=1)
    pdf.ln()
    pdf.set_font("DejaVu", size=14, style="B")
    pdf.cell(200, 6, txt6, ln=1)
    pdf.set_font("DejaVu", size=13)

    # вывод верхней таблицы со счётом за квартал
    data_z = ['Denumirea', 'Pret, lei']
    spacing = 1
    col_width = pdf.w / 2.22  # 4.5
    row_height = pdf.font_size * 1.2
    # Zagolovok
    pdf.set_font("DejaVu", size=13, style="B")
    for item in data_z:
        pdf.cell(col_width, row_height * spacing, txt=item, border=1, align="C")
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=13)
    # print("-" * 100)
    for row in data:
        # print("ROW====", row[0], row[1])
        pdf.cell(col_width, row_height * spacing, txt=str(row[0]), border=1)
        pdf.cell(col_width, row_height * spacing, txt=f"{float(row[1]):.2f}",
                 border=1,
                 align="R")
        pdf.ln(row_height * spacing)

    pdf.ln()
    pdf.cell(200, 6, txt_sign1, ln=1)
    pdf.cell(200, 10, txt_sign2, ln=1)
    pdf.ln()
    pdf.set_font("DejaVu", size=10)
    pdf.set_text_color(220, 50, 50)
    pdf.multi_cell(0, 4, txt_ps2)
    pdf.multi_cell(0, 4, txt_ps3)
    pdf.image('/Users/nd/Desktop/baza/stamp.png', x=xS, y=67 + yS * 6, w=38)
    pdf.image('/Users/nd/Desktop/baza/sign.png', x=xS - 25, y=80 + yS * 6, w=13)

    if year:
        # вывод счёта за год
        pdf.ln()
        pdf.multi_cell(0, 4, txt_ps1)
        # pdf.ln()
        pdf.multi_cell(0, 4, txt_ps4)
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("DejaVu", size=20, style="B")
        pdf.cell(200, 10, txt1, ln=1, align="C")
        pdf.set_font("DejaVu", size=13)
        pdf.cell(200, 6, txt2, ln=1)
        pdf.cell(200, 6, txt3, ln=1)
        pdf.cell(200, 6, txt4, ln=1)
        pdf.cell(200, 6, txt5, ln=1)
        pdf.ln()

        pdf.set_font("DejaVu", size=13, style="B")
        pdf.cell(200, 6, txt6, ln=1)
        pdf.set_font("DejaVu", size=13, style="B")
        for item in data_z:
            pdf.cell(col_width, row_height * spacing, txt=item, border=1,
                     align="C")
        pdf.ln(row_height * spacing)

        pdf.set_font("DejaVu", size=13)
        pdf.cell(col_width, row_height * spacing,
                 txt="Servicii_de_registrator_Anul_2022", border=1)
        pdf.cell(col_width, row_height * spacing,
                 txt=f"{float(table.loc[p1][49]):.2f}", border=1,
                 align="R")
        pdf.ln(row_height * spacing)

        pdf.ln()
        pdf.cell(200, 6, txt_sign1, ln=1)
        pdf.cell(200, 10, txt_sign2, ln=1)
        pdf.image('/Users/nd/Desktop/baza/stamp.png', x=xS, y=201 + yS * 6,
                  w=38)
        pdf.image('/Users/nd/Desktop/baza/sign.png', x=xS - 25, y=213 + yS * 6,
                  w=13)

    emitent = table.loc[p1][1]
    emitent = emitent.replace('"', "")
    emitent = emitent.replace(',', "")
    emitent = emitent.replace(' ', "_")
    emitent = emitent.replace("Ţ", "T")
    emitent = emitent.replace("Ș", "S")
    emitent = emitent.replace("Ş", "S")
    emitent = emitent.replace("Ă", "A")
    emitent = emitent.replace("Î", "I")
    emitent = emitent.replace("Â", "I")

    path_file = f"/Users/nd/Desktop/baza/Conturi/Cont_{emitent}.pdf"
    try:
        print("Saving a file... ", path_file, end="\t")
        pdf.output(path_file)
        print("Saved!")
    except UnicodeEncodeError:
        print("\nBUG! File not saved!", table.loc[p1][1], path_file)

    if send_email:
        status = True
        d1 = time.strftime("%d.%m.%Y")
        t1 = time.strftime("%H:%M:%S")
        if str(p2) == 'nan':
            print("Verify email - email not found!!!")
            status = False
            tmp_status = "no email"
        # print([emitent, p2, d1, t1, status])
        # можно поменять NaN на что-то другое
        # tmp_status = "OK" if status else " "
        # report.append([emitent, str(p2), d1, t1, tmp_status])

        if status:
            email_sender = "grupa_financiara@mail.ru"
            email_password = "p0KaR11DGgTDCNycPfQs"  # нужно сгенерировать новый пароль https://id.mail.ru/security
            email_receiver = p2

            body = """
            ВНИМАНИЕ!!! Налоговая накладная на наши услуги отписывается посредством системы "E-FACTURA" по длинному циклу. Просим своевременно подписывать её. Все вопросы и пожелания принимаются по телефону 022-27-23-13 (бухгалтерия).
            IN ATENTIA CONTABILITATII!!! Factura fiscala pentru serviciile acordate SA "Grupa Financiara" este perfectata prin intermediul "E-FACTURA" - ciclu mare. Rugam semnarea acestei facturi in timp util. Ralatii la telefon 022-27-37-13 (contabilitatea).
    
            Acest document,poate conține date personale. Dacă l-ați primit 
            din greșeală: nu aveţi dreptul de divulgare, păstrare, transmitere a acestor date. Vă rugăm să anunţaţi imediat Grupa Financiară pe e-mail, adresa juridică şi/sau la tel. 022 27-18-45.
            """
            em = MIMEMultipart()
            em['From'] = email_sender
            em['To'] = email_receiver
            em['subject'] = "Cont pentru achitarea serviciilor SA Grupa " \
                            "Financiara (SA " + table.loc[p1][1] + ")"
            em.attach(MIMEText(body, 'plain'))

            filename = "Cont_" + emitent + ".pdf"
            fn = "/Users/nd/Desktop/baza/Conturi/" + filename
            # attachment = open(fn, "rb")
            attachment = open(path_file, "rb")
            part = MIMEBase('application', 'octet-stream')  #----------------
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition',
                            "attachment; filename= %s" % filename)

            em.attach(part)
            try:
                print("path_file:", path_file)
                print("email_receiver:", p2)
                with smtplib.SMTP_SSL('smtp.mail.ru', 465) as smtp:
                    smtp.login(email_sender, email_password)
                    smtp.sendmail(email_sender, email_receiver.split(", "),
                                  em.as_string())
                print("Email send")
                tmp_status = "Send email"
            except:
                print("Dont send")
                tmp_status = "email not sent"
        report.append([emitent, str(p2), d1, t1, tmp_status])
        f = open('report.txt', "a")
        f.write(f"{emitent}, {str(p2)}, {d1}, {t1}, {tmp_status}\n")
        f.close()


def save_report():  # Создаём и сохраняем на диске
    # --------------------------------------------
    pdf = FPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.add_font('DejaVu', '',
                 '/Users/nd/Desktop/baza/fonts/DejaVuSansCondensed.ttf',
                 uni=True)
    pdf.add_font('DejaVu', 'B',
                 '/Users/nd/Desktop/baza/fonts/DejaVuSans-Bold.ttf',
                 uni=True)

    pdf.set_font("DejaVu", size=10, style="B")

    # вывод верхней таблицы со счётом за квартал
    dat = ['Denumirea', 'email', 'Data', 'Timp', 'Send']
    spacing = 1
    # col_width = pdf.w / 2.22  # 4.5
    row_height = pdf.font_size * 1.2
    # Zagolovok
    pdf.set_font("DejaVu", size=11, style="B")
    r1 = [90, 130, 22, 18, 18]
    s = 0
    for item in dat:
        pdf.cell(r1[s], row_height * spacing, txt=item, border=1, align="C")
        s += 1
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=9)
    for row in report:
        pdf.cell(r1[0], row_height * spacing, txt=row[0], border=1)
        pdf.cell(r1[1], row_height * spacing, txt=row[1], border=1, align="R")
        pdf.cell(r1[2], row_height * spacing, txt=row[2], border=1, align="R")
        pdf.cell(r1[3], row_height * spacing, txt=row[3], border=1, align="R")
        pdf.cell(r1[4], row_height * spacing, txt=str(row[4]), border=1,
                 align="R")
        pdf.ln(row_height * spacing)

    pdf.ln()
    # print(report)
    pdf.output('/Users/nd/Desktop/baza/Conturi/report.pdf')
    print("Report saved")
    # --------------------------------------------


def cont_for_one():
    emitent_year = False
    emitent_name = input("Наименование АО: ").upper()
    trim = input(
        "Счёт за какой период (1-trim.I, 2-trim.II, 3-trim.III, 4-trim.IV, "
        "0-весь год): ")
    if trim in ["1", "2", "3"]:
        emitent_year = True if input(
            "Нужен ли счёт за год? 0-нет, 1-да: ") == "1" else False
    send_email = True if input(
        "0-Создать счёт, 1-Создать и отправить счёт: ") == "1" else False
    result = search_SA(emitent_name)
    if result:
        # print(result[0], result[1])
        create_pdf(result[0], result[1], trim, emitent_year, send_email)
        if send_email:
            print(report)
    else:
        print("Нет такого АО")


def conts_for_all():
    emitent_year = False
    trim = int(input(
        "Счёт за какой период (1-trim.I, 2-trim.II, 3-trim.III, 4-trim.IV, "
        "0-весь год): "))
    if trim in [1, 2, 3]:
        emitent_year = True if input(
            "Нужен ли счёт за год? 0-нет, 1-да: ") == "1" else False
    send_email = True if input(
        "0-Создать счёта, 1-Создать и отправить счёта: ") == "1" else False

    for i in range(2, len(table) - 1):
        trim_column = [49, 16, 27, 38, 49]  # Номера столбцов с задолженностями
        if str(table.loc[i][2]) != "STOP" and str(table.loc[i][2]) != "EXPIRAT":
            if str(table.loc[i][trim_column[trim]]) != 'nan' and int(
                    table.loc[i][trim_column[trim]]) > 0:
                result = search_SA(str(table.loc[i][1]))
                # print(result)
                create_pdf(result[0], result[1], trim, emitent_year, send_email)
    save_report()


def infos():
    list_for_send = []
    list_email_for_send = []
    trim = 3
    for i in range(2, len(table) - 1):
        trim_column = [49, 16, 27, 38, 49]  # Номера столбцов с задолженностями
        if str(table.loc[i][2]) != "STOP" and str(table.loc[i][2]) != "EXPIRAT":
            if str(table.loc[i][trim_column[trim]]) != 'nan' and int(
                    table.loc[i][trim_column[trim]]) > 0:
                result = search_SA(str(table.loc[i][1]))
                list_for_send.append(str(table.loc[result[0]][1]))
                list_email_for_send.append(result[1])
    print("Проверка! Всего SA начасление > 0 ", len(list_for_send),
          ". Всего Email", len(list_email_for_send))
    for i in range(len(list_for_send)):
        print(f"{list_for_send[i]:>50} ---> {list_email_for_send[i]}")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    main()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
