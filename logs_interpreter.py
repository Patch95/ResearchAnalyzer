# -*- coding: utf-8 -*-

from loginfo import LogInfo
import os
import scriptchecker
import xlsxwriter
from constants_analysis_value import *


path = os.getcwd() + "\logs_to_treat\\"

logsList = os.listdir(path)

listLogInfo = []


# Function that checks if there are folders or not in the address logs_to_treat
# Check the logs of each folder with the scriptchecker script
# For each folder generates a loginfo instance and creates a list
def load_loginfo():
    for logfolder in logsList:
        click_images = os.path.join(path, logfolder+"\\click_images\\")
        detailed_log = os.path.join(path, logfolder+"\\detailed_log\\")
        system_log = os.path.join(path, logfolder + "\\system_log\\")
        timed_screenshot = os.path.join(path, logfolder + "\\timed_screenshots\\")

        click_images_files = [f for f in os.listdir(click_images) if f.endswith('.txt')]
        if click_images_files.__len__() >= 1:
            click_fn = os.path.join(click_images, click_images_files[0])
        else:
            print("Failed to locate the file click_images")
            print click_images
            exit()

        detailed_log_files = [f for f in os.listdir(detailed_log) if f.endswith('.txt')]
        if detailed_log_files.__len__() >= 1:
            keys_fn = os.path.join(detailed_log, detailed_log_files[0])
        else:
            print("Failed to locate the file detailed_log")
            exit()

        scriptchecker.arranger(keys_fn, click_fn)
        keys_fn_new = keys_fn.replace(".txt", "_new.txt")
        click_fn_new = click_fn.replace(".txt", "_new.txt")

        system_log_files = [f for f in os.listdir(system_log) if f.endswith('.txt')]
        if system_log_files.__len__() >= 1:
            system_fn = os.path.join(system_log, system_log_files[0])
        else:
            print("Failed to locate the file system_log")
            exit()

        timed_screenshot_files = [f for f in os.listdir(timed_screenshot) if f.endswith('.txt')]
        if timed_screenshot_files.__len__() >= 1:
            screenshoots_fn = os.path.join(timed_screenshot, timed_screenshot_files[0])
        else:
            print("Failed to locate the file timed_screenshot")
            exit()

        li = LogInfo(click_fn_new,  # Click data file here
                     keys_fn_new,  # Detailed log file here
                     screenshoots_fn,  # Timed screenshot log file here
                     system_fn,  # System log data here"""
                     )

        listLogInfo.append(li)


# Function that generates the total time of the traduction session by subtraction
def total_session_time():

    for li in listLogInfo:
        print "h"
        mp = li.get_mixed_parser()
        merged = mp.get_merged()
        last = False
        for i in range(0, merged.__len__()):
            line = merged[i]
            event = str(line).split(",")

            if event.__len__() == 12:
                # date, real_time, program_name, username, window_id, window_title, ms, interruption, x, y,
                _, _, _, _, _, _, ms, _, _, _, _, _ = line

            else:
                _, _, _, _, _, _, ms, _, _, _, _ = line

            if i == 0:
                init = ms
            elif i == merged.__len__() - 1:
                last = ms

        print "total session time: "
        print int(last) - int(init)
        print "time by active active window"
        print li.get_time_by_active_window()
        print


# Function that create a string from the list elements
def get_info_of_list(dataList):
    if dataList.__len__() == 0:
        return ""
    else:
        strtemp = str(dataList[0])
        if strtemp.__contains__(".exe") or strtemp.__contains__(".EXE"):
            strtemp = generate_key_name(str(strtemp))
        for x in range(1, dataList.__len__()):
            data = dataList[x]
            if str(data).__contains__(".exe") or str(data).__contains__(".EXE"):
                strtemp += ", " + generate_key_name(str(data))
            else:
                strtemp += ", " + str(data)

        return strtemp


# Main function that creates a worksheet for the clicks interruptions
def generate_click_info(li, workbook, addname):

    # Add a bold format to use to highlight cells.e
    bold = workbook.add_format({'bold': True})
    myformat = workbook.add_format({'bold': True})
    myformat.set_align('center')
    myformat.set_align('vcenter')

    name = "ClicksInfo_" + addname

    worksheet2 = workbook.add_worksheet(name)
    worksheet2.set_column('A:B', 35)
    worksheet2.set_column('E:F', 15)
    worksheet2.set_column('H:I', 25)

    click_sumary = li.print_click_summary()
    worksheet2.merge_range("A1:B1", "Resumen de Clicks", myformat)
    worksheet2.write("A2", "Cantidad total de clicks:", bold)
    worksheet2.write("A3", "Ventanas en las cuales se hacen clicks:", bold)
    worksheet2.write("A4", "Tipos de clicks usados:", bold)
    worksheet2.write("A5", "Tiempo de presión de click promedio:", bold)
    worksheet2.write("A6", "Varianza en la presión de click:", bold)
    worksheet2.write("A7", "Desviación estándar en la presión de click:", bold)

    worksheet2.write("B2", int(click_sumary[0]))
    strV = get_info_of_list(click_sumary[1])
    worksheet2.write("B3", strV)

    strT = get_info_of_list(click_sumary[2])
    worksheet2.write("B4", strT)

    worksheet2.write("B5", float(click_sumary[3]))
    worksheet2.write("B6", float(click_sumary[4]))
    worksheet2.write("B7", float(click_sumary[5]))

    # Write some simple text.
    worksheet2.write('C10', 'X', myformat)
    worksheet2.write('D10', 'Y', myformat)
    worksheet2.write('E10', 'Resolucion', myformat)
    worksheet2.write('F10', 'Tipo de Click', myformat)
    worksheet2.write('G10', 'MSG', myformat)
    worksheet2.write('H10', 'Nombre de programa', myformat)
    worksheet2.write('I10', 'Titulo', myformat)

    up = 0
    all_pressed_clicks = li.get_all_pressed_clicks()
    for x in range(0, all_pressed_clicks.__len__()):
        pressed_click = all_pressed_clicks[x]
        ind = str(11 + x - up)
        if pressed_click[4] == "down":
            worksheet2.write("C" + ind, int(pressed_click[0]))
            worksheet2.write("D" + ind, int(pressed_click[1]))
            worksheet2.write("E" + ind, pressed_click[2])
            worksheet2.write("F" + ind, pressed_click[3])
            worksheet2.write("G" + ind, pressed_click[4])
            if pressed_click[5].__contains__(".exe") or pressed_click[5].__contains__(".EXE"):
                worksheet2.write("H" + ind, generate_key_name(pressed_click[5]))
            else:
                worksheet2.write("H" + ind, pressed_click[5])
            worksheet2.write("I" + ind, pressed_click[6])
        else:
            up += 1


# get combos and filter per dowm
def get_combos(all_keys):
    # CTRL + KEY combos
    # Eg. CTRL + V, CTRL + C, etc
    new_text = []

    combo = ""
    combing = False
    found_key = None
    SPECIAL = {'ctrll', 'ctrlr', 'altl', 'altr', 'lshift', 'rshift'}
    for k in range(0, all_keys.__len__()):
        key = all_keys[k]
        keystroke = key[2]
        event_type = key[3]

        if (event_type == 'down') and (not combing) and (keystroke in SPECIAL):
            combo += keystroke
            found_key = keystroke
            combing = True
        elif event_type == 'down' and keystroke != combo and combing:
            combo += "+"+keystroke
        elif event_type == 'down' and keystroke == combo and combing:
            pass
        elif event_type == 'up' and combing and keystroke == found_key:
            new_text.append(combo)
            combing = False
            combo = ""
            found_key = None
        else:
            if event_type == 'down':
                new_text.append(keystroke)

    return new_text


def found_text_repetitions(l_changed_text):
    if l_changed_text.__len__() != 0:
        new_list = []
        key = l_changed_text[0]
        count = 1
        for x in range(1, l_changed_text.__len__()):

            if key == l_changed_text[x]:
                count += 1
            elif count > 1:
                new_list.append(str(count)+key)
                count = 1
            else:
                new_list.append(key)

            key = l_changed_text[x]
        if count > 1:
            new_list.append(str(count) + key)
        else:
            new_list.append(key)
        return new_list
    else:
        return []


# Function that create a partial reconstruction of the text
def generate_words_reconstruction(all_keys, workbook, addname):
    keys_per_title = []
    temp_list = [all_keys[0]]
    i = 1
    all_keys.__delitem__(0)
    for pair in all_keys:
        if temp_list[i-1][1] == pair[1]:
            temp_list.append(pair)
            i += 1
        else:
            keys_per_title.append(temp_list)
            temp_list = [pair]
            i = 1
    keys_per_title.append(temp_list)

    bold = workbook.add_format({'bold': True})
    myformat = workbook.add_format({'bold': True})
    myformat.set_align('center')
    myformat.set_align('vcenter')
    wrapformat = workbook.add_format()
    wrapformat.set_text_wrap()

    name = "Textinfo_" + addname
    worksheet = workbook.add_worksheet(name)

    worksheet.set_column('A:C', 30)
    worksheet.merge_range("A1:C1", "Teclas por ventana", myformat)

    worksheet.set_column('H:H', 40)
    worksheet.write("H1", "Filtrado de busquedas", myformat)

    index = 3
    i = 0
    search_index = 2
    # genarate text per down
    for lot in keys_per_title:
        original_text = ""
        resourse = lot[0][0]
        tool = lot[0][1]
        worksheet.write("A" + str(index + (11 * i)), "Recurso y Herramienta:", bold)
        worksheet.write("B" + str(index + (11 * i)), resourse)  # Recurso excel
        worksheet.write("C" + str(index + (11 * i)), tool)  # Herramienta excel

        if tool.__contains__("Buscar con Google - Google Chrome"):
            busqueda = tool.replace(" - Buscar con Google - Google Chrome", " ")
            worksheet.write("H" + str(search_index), busqueda)  # Recurso excel
            search_index += 1

        for k in range(0, lot.__len__()):
            key = lot[k]
            if key[3] != 'up':
                if key[2].__len__() == 1:
                    original_text += key[2]
                else:
                    original_text += " \"" + key[2] + "\" "
        worksheet.merge_range("A" + str((index + 1) + (11 * i)) + ":C" + str((index + 4) + (11 * i)), original_text, wrapformat)
        i += 1

    index = 3
    i = 0
    # generate modificated text
    for lot in keys_per_title:
        l_changed_text = get_combos(lot)
        changed_text = ""

        # replace interruptions with understandable name
        for x in range(0, l_changed_text.__len__()):
            key = l_changed_text[x]
            try:
                l_changed_text[x] = REPLACE_KEYS[key]
            except:
                l_changed_text[x] = key

        l_changed_text = found_text_repetitions(l_changed_text)  # call to funtion

        # create string for the excel file
        for x in range(0, l_changed_text.__len__()):
            key = l_changed_text[x]
            if key.__len__() > 1:
                changed_text += " \"" + key + "\""
            elif x != 0 and l_changed_text[x-1].__len__() > 1:
                changed_text += " " + key
            else:
                changed_text += key

        worksheet.merge_range("A" + str((index + 6) + (11 * i)) + ":C"
                              + str((index + 9) + (11 * i)), changed_text, wrapformat)
        i += 1


# Main function that creates a worksheet for the key interruptions
# call generate_words_reconstruction and compilation_attempts functions
def generate_key_info(li, workbook, addname):

    # Add a bold format to use to highlight cells.e
    bold = workbook.add_format({'bold': True})
    myformat = workbook.add_format({'bold': True})
    myformat.set_align('center')
    myformat.set_align('vcenter')

    name = "KeysInfo_" + addname
    worksheet3 = workbook.add_worksheet(name)
    worksheet3.set_column('A:B', 35)
    worksheet3.set_column('C:D', 15)
    key_sumary = li.print_key_summary()

    worksheet3.merge_range("A1:B1", "Resumen de Teclas", myformat)
    worksheet3.write("A2", "Cantidad de teclas presionadas:", bold)
    worksheet3.write("A3", "Teclas únicas presionadas:", bold)
    worksheet3.write("A4", "Ventanas donde se esribió:", bold)
    worksheet3.write("A5", "Tiempo de presión de tecla promedio:", bold)
    worksheet3.write("A6", "Varianza de presión de tecla:", bold)
    worksheet3.write("A7", "Desviación de presión de tecla:", bold)
    worksheet3.write("A8", "Teclas de función usadas:", bold)
    worksheet3.write("A9", "Patrones de borrado usados:", bold)
    worksheet3.write("A10", "Teclas de movimiento usadas:", bold)
    worksheet3.write("A11", "Combos de teclas usados:", bold)

    worksheet3.write("B2", int(key_sumary[0]))
    strTemp = get_info_of_list(key_sumary[1])
    worksheet3.write("B3", strTemp)
    strTemp = get_info_of_list(key_sumary[2])
    worksheet3.write("B4", strTemp)
    worksheet3.write("B5", float(key_sumary[3]))
    worksheet3.write("B6", float(key_sumary[4]))
    worksheet3.write("B7", float(key_sumary[5]))
    strTemp = get_info_of_list(key_sumary[6])
    worksheet3.write("B8", strTemp)
    strTemp = get_info_of_list(key_sumary[7])
    worksheet3.write("B9", strTemp)
    strTemp = get_info_of_list(key_sumary[8])
    worksheet3.write("B10", strTemp)

    # eliminat ctrl repetition and filter combos
    templList = []
    if key_sumary[9].__len__() == 0:
        strTemp = ""
    else:
        strTemp = key_sumary[9][0][0] + "+" + key_sumary[9][0][-1]
        templList.append(strTemp)
        for x in range(1, key_sumary[9].__len__()):
            key_list = key_sumary[9][x]
            combo = key_list[0] + "+" + key_list[-1]
            if not (combo in templList):
                strTemp += ", " + combo
                templList.append(combo)

    worksheet3.write("B11", strTemp)

    worksheet3.merge_range("A13:D13", "Teclas resumen total", myformat)
    worksheet3.write("A14", "Ventana", bold)
    worksheet3.write("B14", "Titulo de Ventana", bold)
    worksheet3.write("C14", "Tecla presionada", bold)
    worksheet3.write("D14", "Tipo de Evento", bold)

    all_pressed_keys = li.get_all_pressed_keys()

    all_keys = []
    for x in range(0, all_pressed_keys.__len__()):
        line = all_pressed_keys[x]
        keystroke, event_type, window, window_title = line
        if window.__contains__(".exe") or window.__contains__(".EXE"):
            worksheet3.write("A" + str(15 + x), generate_key_name(window))
            all_keys.append([generate_key_name(window), window_title, keystroke, event_type])
        else:
            worksheet3.write("A" + str(15 + x), window)
            all_keys.append([window, window_title, keystroke, event_type])
        worksheet3.write("B" + str(15 + x), window_title)
        worksheet3.write("C" + str(15 + x), keystroke)
        worksheet3.write("D" + str(15 + x), event_type)

    generate_words_reconstruction(all_keys, workbook, addname)

    attemps = list(
        filter(lambda (window, window_title, key, event): window == "pythonw" and key == "f5" and event == "up",
               all_keys))
    return attemps.__len__()


# Create chart in excel
def generate_total_plot(workbook, worksheet, users):
    usersamount = users.__len__()

    # Create a new chart object.
    chart = workbook.add_chart({'type': 'column'})

    for i in range(0, usersamount):
        # Add a series to the chart.
        # [sheetname, first_row, first_col, last_row, last_col]
        # columna de categorias y columna de valores
        # categorias F5 =  3, 5 y F6 = 4, 5
        chart.add_series({
            'categories': ['GeneralInfo', 4, 5, 3+2, 5],  # +2 ya que solo hay dos categorias
            'values': ['GeneralInfo', 4, 6+i, 3+2, 6+i],  #
            'name': users[i]
        })

    chart.set_title({
        'name': 'Tiempo total de desarrollo',
    })

    chart.set_y_axis({
        'name': 'Duración Minutos',
    })

    # Insert the chart into the worksheet.
    worksheet.insert_chart(3, usersamount + 7, chart)


# Create chart in excel
def generate_resources_plot(workbook, worksheet, users, resourcesamount):
    usersamount = users.__len__()

    # Create a new chart object.
    chart = workbook.add_chart({'type': 'column'})

    for i in range(0, usersamount):
        # Add a series to the chart.
        #     [sheetname, first_row, first_col, last_row, last_col]

        chart.add_series({
            'categories': ['GeneralInfo', 22, 5, 21+resourcesamount, 5],
            'values': ['GeneralInfo', 22, 6+i, 21+resourcesamount, 6+i],
            'name': users[i]
        })

    chart.set_title({
        'name': 'Distribución de tiempo en recursos principales',
    })

    chart.set_y_axis({
        'name': 'Duración Minutos',
    })

    # Insert the chart into the worksheet.
    worksheet.insert_chart(22, usersamount+7, chart)


# Function that eliminate excel extension
def generate_key_name(key):
    key_name = ""
    indice = key.__len__()-5
    while(indice != 0):
        if key[indice] == "\\":
            break
        else:
            key_name += key[indice]
            indice -= 1

    return key_name[::-1]


# Main function that create excel file
def generate_full_data():

    workbook = xlsxwriter.Workbook("United_Analysis.xlsx")
    worksheet = workbook.add_worksheet("GeneralInfo")
    init = 3

    # formast to use
    myformat = workbook.add_format({'bold': True})
    myformat.set_align('center')
    myformat.set_align('vcenter')
    subjectformat = workbook.add_format({'bold': True})
    subjectformat.set_align('center')
    subjectformat.set_align('vcenter')

    # Widen the first column to make the text clearer.
    worksheet.set_column('B:B', 20)
    worksheet.set_column('C:C', 30)
    worksheet.set_column('D:D', 12)
    worksheet.set_column('E:E', 3)
    worksheet.set_column('F:G', 22)

    # Write some simple text.
    worksheet.write("A" + str(init), 'Sujeto', myformat)
    worksheet.write("B" + str(init), "Herramientas", myformat)
    worksheet.write("C" + str(init), "Milisegundos \n por recurso", myformat)
    worksheet.write("D" + str(init), '%Recurso', myformat)

    indtotals = [[], []]
    resources = []
    users = []
    compilation_attemps = []
    for i in range(0, listLogInfo.__len__()):
        li = listLogInfo[i]

        keys = li.get_time_by_active_window().keys()

        t = 0
        for j in range(0, keys.__len__()):
            key = keys[j]
            if key == "total":
                t = 1
            else:
                ind1 = "B" + str(j + (init+1) - t)
                if key.__contains__(".exe") or key.__contains__(".EXE"):
                    worksheet.write(ind1, generate_key_name(key))
                else:
                    worksheet.write(ind1, key)

                ind2 = "C" + str(j + (init+1) - t)
                worksheet.write(ind2, li.get_time_by_active_window()[key])

                if not resources.__contains__(key):
                    resources.append(key)

        ind = "A" + str(init+1) + ":" + "A" + str(j + (init+1))

        logname = logsList[i].split("-")

        subjectformat = workbook.add_format({'bold': True})
        subjectformat.set_bg_color("#81DAF5")
        subjectformat.set_align('center')
        subjectformat.set_align('vcenter')
        worksheet.merge_range(ind, logname[1], subjectformat)

        worksheet.write(21, 6+i, logname[1])  # write users names for chart
        users.append(logname[1])

        ind1 = "B" + str(j + (init+1))
        worksheet.write(ind1, "total")
        ind2 = "C" + str(j + (init+1))

        worksheet.write(ind2, li.get_time_by_active_window()["total"])

        indtotals[0].append(logname[1])
        indtotals[1].append(li.get_time_by_active_window()["total"])

        actual_user = logname[1] + "_" + logname[2]

        # call to funtions
        generate_click_info(li, workbook, actual_user)

        # function returns number of compilations of user
        compilation_attemps.append(generate_key_info(li, workbook, actual_user))

        # percent of total sesion time
        percent_fmt = workbook.add_format({'num_format': '0.00%'})
        for lin in range(0, keys.__len__()):
            cell = "D" + str(init+1+lin) + ":D" + str(init+1+lin)
            formula = "{=SUM(C" + str(init+1+lin) + "/C" + str(j + (init+1)) + ")}"
            worksheet.write_array_formula(cell, formula, percent_fmt)

        init = ((j + (init + 1)) + 1)

    worksheet.merge_range(2, 5, 2, 5+users.__len__(), "Desarrollo del usuario", subjectformat)
    worksheet.merge_range(20, 5, 20, 5+users.__len__(), "Distribución de tiempo en recursos principales", subjectformat)

    # Write resource for chart
    names_resources = []
    for nkey in resources:
        if nkey.__contains__(".exe") or nkey.__contains__(".EXE"):
            names_resources.append(generate_key_name(nkey))
        else:
            names_resources.append(nkey)
    worksheet.write_column('F23', names_resources)

    for nli in range(0, listLogInfo.__len__()):
        lli = listLogInfo[nli]
        for r in range(0, resources.__len__()):
            resource = resources[r]
            try:
                mlsec = lli.get_time_by_active_window()[resource]
                mnsec = float(mlsec) / 60000.0
                worksheet.write(22 + r, 6 + nli, mnsec)
            except:
                worksheet.write(22 + r, 6 + nli, 0.0)

    users = indtotals[0]
    times = indtotals[1]

    for i in range(times.__len__()):
        times[i] = times[i] / 60000.0

    #worksheet.write_column('F4', users)
    #worksheet.write_column('G4', times)

    worksheet.write_row('G4', users)
    worksheet.write_row('G5', times)
    worksheet.write_row('G6', compilation_attemps)
    worksheet.write("F5:F5", "Tiempo total de sesión")
    worksheet.write("F6:F6", "Número de compilaciones")

    generate_total_plot(workbook, worksheet, users)
    generate_resources_plot(workbook, worksheet, users, resources.__len__())

    print "Excel file writed"


# Function that creates a single excel per folder
def generate_separate_data():

    for i in range(0, listLogInfo.__len__()):
        li = listLogInfo[i]

        logname = logsList[i].split("-")
        name = "Analysis_" + logname[1] + "_" + logname[2] + ".xlsx"
        workbook = xlsxwriter.Workbook(name)
        worksheet = workbook.add_worksheet("GeneralInfo")

        # Add a bold format to use to highlight cells.e
        bold = workbook.add_format({'bold': True})
        myformat = workbook.add_format({'bold': True})
        myformat.set_align('center')
        myformat.set_align('vcenter')

        # Widen the first column to make the text clearer.

        worksheet.set_column('C:D', 30)
        #worksheet.set_row(2, 40)

        # Write some simple text.
        worksheet.write('B3', 'Sujeto', myformat)
        #worksheet.write('B4', logname[1], myformat)
        worksheet.write('C3', "Herramientas", myformat)
        worksheet.write('D3', "Milisegundos \n por recurso", myformat)

        #print "time by active active window"
        #print li.get_time_by_active_window()
        keys = li.get_time_by_active_window().keys()
        t = 0
        for j in range(0, keys.__len__()):
            key = keys[j]
            if key == "total":
                t = 1
            else:
                ind1 = "C" + str(j+4-t)
                worksheet.write(ind1, key)
                ind2 = "D" + str(j+4-t)
                worksheet.write(ind2, li.get_time_by_active_window()[key])
                #print key, li.get_time_by_active_window()[key]

        ind = "B4:" + "B" + str(j+4)
        worksheet.merge_range(ind, logname[1], myformat)

        ind1 = "C" + str(j+4)
        worksheet.write(ind1, "total")
        ind2 = "D" + str(j+4)
        worksheet.write(ind2, li.get_time_by_active_window()["total"])

        generate_click_info(li, workbook, "")

        generate_key_info(li, workbook, "")

    print("Excel file writed")


def main():
    load_loginfo()
    generate_full_data()
    #generate_separate_data()


if __name__ == "__main__":
    main()



