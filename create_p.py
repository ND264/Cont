def create_pdf(p1, p2, trim, year):
    '''
    Функция создающая pdf-документ и записывает файл
    :param p1: Номер строки где находится АО на листе "2022"
    :param p2: Номер строки где находится АО на листе "data"
    :param trim: номер квартала 1-4 или 0 если счёт на год
    :param year: true/false если нужен 2-й счет до конца года
    :return:
    '''
    period = [{'Anul_2022': 49},
              {'Datorie': 5, 'Adonament_trim.1': 6, 'Extras': 7,
               'Lista': 8, 'Scrisori': 9, 'Emisie': 10, 'Informatie': 11,
               'Consalting': 12, 'Alte': 13, 'Spre_achitare': 16},
              {'Datorie': 16, 'Adonament_trim.2': 17, 'Extras': 18,
               'Lista': 19, 'Scrisori': 20, 'Emisie': 21, 'Informatie': 22,
               'Consalting': 23, 'Alte': 24, 'Spre_achitare': 27},
              {'Datorie': 27, 'Adonament_trim.3': 28, 'Extras': 29,
               'Lista': 30, 'Scrisori': 31, 'Emisie': 32,
               'Informatie': 33, 'Consalting': 34, 'Alte': 35,
               'Spre_achitare': 38},
              {'Datorie': 38, 'Adonament_trim.4': 39, 'Extras': 40,
               'Lista': 41, 'Scrisori': 42, 'Emisie': 43,
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
            print("key:", k, "value:", table.loc[p1][v])
    print("DATA: ", data)

    date = time.strftime("%d.%m.%Y")  # "30.08.2022"
    txt1 = f"FACTURA din {date}"
    txt2 = u'Furnizor: S.A. "Grupa Financiară", IDNO: 1002600054323'
    txt3 = 'Adresa: MD2001, mun. Chişinău, str. A.Bernardazzi, 7'
    txt4 = 'IBAN: MD17AG000000022515490120 la B.C. "MOLDOVA-AGROINDBANK" S.A. s. Tighina'
    txt5 = 'Cod bancar: AGRNMD2X472'
    txt6 = u'Platitor: S.A. "' + table.loc[p1][1] + '"'
    txtPS1 = u"""Если у Вашей компании есть возможность оплатить услуги за оставшийся период 2022 года, ниже представлен счёт по которому можно произвести платёж."""
    txtPS2 = u"""ВНИМАНИЕ!!! Налоговая накладная на наши услуги отписывается посредством системы 'E-FACTURA' по длинному циклу. Просим своевременно подписывать её. Все вопросы и пожелания принимаются по телефону 022-27-23-13 (бухгалтерия)."""
    txtPS3 = u"""IN ATENTIA CONTABILITATII!!! Factura fiscala pentru serviciile acordate SA 'Grupa Financiară' este perfectata prin intermediul 'E-FACTURA' - ciclu mare. Rugam semnarea acestei facturi in timp util. Ralatii la telefon 022-27-37-13 (contabilitatea)."""
    txt_sign1 = 'Director Viorica BONDAREV'
    txt_sign2 = u'Contabil-şef Maia CIULCOVA'

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
    dataZ = ['Denumirea', 'Pret, lei']
    spacing = 1
    col_width = pdf.w / 2.22  # 4.5
    row_height = pdf.font_size * 1.2
    # Zagolovok
    pdf.set_font("DejaVu", size=13, style="B")
    for item in dataZ:
        pdf.cell(col_width, row_height * spacing, txt=item, border=1, align="C")
    pdf.ln(row_height * spacing)

    pdf.set_font("DejaVu", size=13)
    print("-"*100)
    for row in data:
        print("ROW====", row[0], row[1])
        pdf.cell(col_width, row_height * spacing, txt=str(row[0]), border=1)
        pdf.cell(col_width, row_height * spacing, txt=str(row[1]), border=1,
                 align="R")
        pdf.ln(row_height * spacing)

    pdf.ln()
    pdf.cell(200, 6, txt_sign1, ln=1)
    pdf.cell(200, 10, txt_sign2, ln=1)
    pdf.ln()
    pdf.set_font("DejaVu", size=10)
    pdf.set_text_color(220, 50, 50)
    pdf.multi_cell(0, 4, txtPS2)
    pdf.multi_cell(0, 4, txtPS3)
    pdf.image('/Users/nd/Desktop/baza/stamp.png', x=xS, y=68 + yS * 6, w=38)
    pdf.image('/Users/nd/Desktop/baza/sign.png', x=xS - 30, y=82 + yS * 6, w=13)

    if year:
        # вывод счёта за год
        pdf.ln()
        pdf.multi_cell(0, 4, txtPS1)
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
        for item in dataZ:
            pdf.cell(col_width, row_height * spacing, txt=item, border=1,
                     align="C")
        pdf.ln(row_height * spacing)

        pdf.set_font("DejaVu", size=13)
        pdf.cell(col_width, row_height * spacing, txt="Anul 2022", border=1)
        pdf.cell(col_width, row_height * spacing,
                 txt=str(table.loc[p1][49]), border=1,
                 align="R")
        pdf.ln(row_height * spacing)

        pdf.ln()
        pdf.cell(200, 6, txt_sign1, ln=1)
        pdf.cell(200, 10, txt_sign2, ln=1)
        pdf.image('/Users/nd/Desktop/baza/stamp.png', x=xS, y=199 + yS * 6,
                  w=38)
        pdf.image('/Users/nd/Desktop/baza/sign.png', x=xS - 30, y=211 + yS * 6,
                  w=13)

    emitent = table.loc[p1][1]
    emitent = emitent.replace('"', "")
    emitent = emitent.replace(',', "")
    emitent = emitent.replace(' ', "_")
    emitent = emitent.replace("Ţ", "T")
    emitent = emitent.replace("Ș", "S")
    emitent = emitent.replace("Ă", "A")
    emitent = emitent.replace("Î", "I")
    emitent = emitent.replace("Â", "I")

    path_file = f"/Users/nd/Desktop/baza/Conturi/Cont_{emitent}.pdf"
    try:
        print("Saving a file... ", path_file, end="\t")
        pdf.output(path_file)
        print("Saved!")
    except(UnicodeEncodeError):
        print("\nBUG! File not saved!", table.loc[p1][1], path_file)


def cont_for_one():
    emitent_year = False
    emitent_name = input("Наименование АО: ")
    trim = input(
        "Счёт за какой период (1-trim.I, 2-trim.II, 3-trim.III, 4-trim.IV, 0-весь год): ")
    if trim in ["1", "2", "3"]:
        emitent_year = True if input(
            "Нужен ли счёт за год? 0-нет, 1-да: ") == "1" else False
    #	print(f"Выбор: SA={emitent_name}, trim={emitent_period}, year={emitent_year}")
    # search_emitent(emitent_name, emitent_period, emitent_year)
