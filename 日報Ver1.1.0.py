#  target_row = edit_start_row
#         if edit_input_row_count > edit_row_count:
#            rows_to_insert = edit_input_row_count - edit_row_count

#            for i in range(rows_to_insert):

#                 report_sheet.insert_rows(edit_last_row)
#                 report_sheet['A'+ str(edit_last_row)].number_format = u'm\月d\日'
#                 for j in report_sheet["A1":"J1000"]:
#                     for l in j:
#                         l.border = border
#                 for j in report_sheet["L1":"M1000"]:
#                     for l in j:
#                         l.border = border
#                 edit_row_count += 1
#                 report_workbook.save(report_workbook_path)
#         elif edit_input_row_count < edit_row_count:
#             rows_to_delete = edit_row_count - edit_input_row_count
#             report_sheet.delete_rows(idx=edit_start_row, amount=rows_to_delete)
#             report_workbook.save(report_workbook_path)

def save_contents(confirmed):
    global target_row
    global pending_rows
    global edit_row_count
    global pending_rows
    border = Border(top=side, bottom=side, left=side, right=side)
    report_workbook = load_required_workbook(report_workbook_path, "日報出力 Excel")
    report_sheet = report_workbook.worksheets[0]
    if edit_mode == 1:
        target_row = edit_start_row
        if edit_input_row_count > edit_row_count:
           rows_to_insert = edit_input_row_count - edit_row_count

           for i in range(rows_to_insert):

                report_sheet.insert_rows(edit_last_row)
                report_sheet['A'+ str(edit_last_row)].number_format = u'm\月d\日'
                for j in report_sheet["A1":"J1000"]:
                    for l in j:
                        l.border = border
                for j in report_sheet["L1":"M1000"]:
                    for l in j:
                        l.border = border
                edit_row_count += 1
        elif edit_input_row_count < edit_row_count:
            rows_to_delete = edit_row_count - edit_input_row_count
            report_sheet.delete_rows(idx=edit_start_row, amount=rows_to_delete)

    for n in range(len(confirmed)):
        b = 1
        for row_a in confirmed[n]:
            report_sheet.cell(row=target_row,column=b).value = row_a
            b+=1
        target_row +=1
    pending_rows = []
    try:
        report_workbook.save(report_workbook_path)
        messagebox.showwarning("確認","登録しました。")
        clear_all_rows()
    except PermissionError:
        messagebox.showwarning("確認","登録ファイルが開かれています。")



def search_materials():
    material_search_listbox.delete(0,END)
    material_search_entry_get = str(material_search_entry.get())
    for i in material_search_values:
        if material_search_entry_get in i:
            material_search_listbox.insert(tk.END,i)

def clear_material_search():
    material_search_listbox.delete(0,END)
    material_search_entry.delete(0,tk.END)



def insert_selected_material():
    global confirmed_material_number_1
    global confirmed_material_number_2
    global confirmed_material_number_3
    global confirmed_material_number_4
    global confirmed_material_number_5
    global confirmed_material_number_6
    global confirmed_material_number_7
    global confirmed_material_number_8
    material_index = material_search_listbox.curselection()[0]

    if checkbutton_var_1.get() == "1":
        f = confirmed_material_number_1 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_1.current(f)

    if checkbutton_var_2.get() == "1":
        f = confirmed_material_number_2 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_2.current(f)

    if checkbutton_var_3.get() == "1":
        f = confirmed_material_number_3 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_3.current(f)

    if checkbutton_var_4.get() == "1":
        f = confirmed_material_number_4 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_4.current(f)

    if checkbutton_var_5.get() == "1":
        f = confirmed_material_number_5 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_5.current(f)

    if checkbutton_var_6.get() == "1":
        f = confirmed_material_number_6 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_6.current(f)

    if checkbutton_var_7.get() == "1":
        f = confirmed_material_number_7 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_7.current(f)

    if checkbutton_var_8.get() == "1":
        f = confirmed_material_number_8 = material_search_listbox.get(material_index)[:7]
        f = material_numbers.index(str(f))
        material_number_combobox_8.current(f)

def find_first_empty_row(column_cells, worksheet):
    cell_check_value = config_sheet['D2'].value
    if cell_check_value == None:
        cell_check_mode = 0
    else:
        cell_check_mode = cell_check_value

    if int(cell_check_mode) == 1:
        j = len(column_cells)
        if j > 100:
            i = worksheet.max_row + 1
        else:
            i = 0
            for cell_i in column_cells:
                i += 1
                if cell_i.value == None:
                    break
    elif int(cell_check_mode) == 0:
        i = 0
        for cell_i in column_cells:
                i += 1
                if cell_i.value == None:
                    break
        if cell_i.value != None:
            i = 0
            for cell_i in column_cells:
                    i += 1
                    if cell_i.value == None:
                        break


    return i

def load_rows_for_edit():
    global edit_start_row
    global edit_row_count
    global edit_last_row
    global edit_mode
    global k
    global confirmed_day_1
    global confirmed_day_2
    global confirmed_day_3
    global confirmed_day_4
    global confirmed_day_5
    global confirmed_day_6
    global confirmed_day_7
    global confirmed_day_8
    global confirmed_material_number_1
    global confirmed_material_number_2
    global confirmed_material_number_3
    global confirmed_material_number_4
    global confirmed_material_number_5
    global confirmed_material_number_6
    global confirmed_material_number_7
    global confirmed_material_number_8
    global confirmed_work_class_num_1
    global confirmed_work_class_num_2
    global confirmed_work_class_num_3
    global confirmed_work_class_num_4
    global confirmed_work_class_num_5
    global confirmed_work_class_num_6
    global confirmed_work_class_num_7
    global confirmed_work_class_num_8
    global confirmed_work_num_1
    global confirmed_work_num_2
    global confirmed_work_num_3
    global confirmed_work_num_4
    global confirmed_work_num_5
    global confirmed_work_num_6
    global confirmed_work_num_7
    global start_time_h1
    global start_time_h2
    global start_time_h3
    global start_time_h4
    global start_time_h5
    global start_time_h6
    global start_time_h7
    global start_time_h8
    global start_time_m1
    global start_time_m2
    global start_time_m3
    global start_time_m4
    global start_time_m5
    global start_time_m6
    global start_time_m7
    global start_time_m8
    global end_time_h1
    global end_time_h2
    global end_time_h3
    global end_time_h4
    global end_time_h5
    global end_time_h6
    global end_time_h7
    global end_time_h8
    global end_time_m1
    global end_time_m2
    global end_time_m3
    global end_time_m4
    global end_time_m5
    global end_time_m6
    global end_time_m7
    global end_time_m8
    global confirmed_start_time_h_1
    global confirmed_start_time_h_2
    global confirmed_start_time_h_3
    global confirmed_start_time_h_4
    global confirmed_start_time_h_5
    global confirmed_start_time_h_6
    global confirmed_start_time_h_7
    global confirmed_start_time_h_8
    global confirmed_start_time_m_1
    global confirmed_start_time_m_2
    global confirmed_start_time_m_3
    global confirmed_start_time_m_4
    global confirmed_start_time_m_5
    global confirmed_start_time_m_6
    global confirmed_start_time_m_7
    global confirmed_start_time_m_8
    global confirmed_end_time_h_1
    global confirmed_end_time_h_2
    global confirmed_end_time_h_3
    global confirmed_end_time_h_4
    global confirmed_end_time_h_5
    global confirmed_end_time_h_6
    global confirmed_end_time_h_7
    global confirmed_end_time_h_8
    global confirmed_end_time_m_1
    global confirmed_end_time_m_2
    global confirmed_end_time_m_3
    global confirmed_end_time_m_4
    global confirmed_end_time_m_5
    global confirmed_end_time_m_6
    global confirmed_end_time_m_7
    global confirmed_end_time_m_8

    edit_report_workbook = load_required_workbook(report_workbook_path, "日報出力 Excel")
    edit_report_sheet = edit_report_workbook.worksheets[0]
    if "confirmed_name" in globals() and "confirmed_month_1" in globals() and "confirmed_day_1" in globals():

        _name_ =  confirmed_name
        _date_ = datetime.datetime(int(current_year),int(confirmed_month_1),int(confirmed_day_1),00,00)
        i = 0
        edit_report_workbook = load_required_workbook(report_workbook_path, "日報出力 Excel")
        edit_report_sheet = edit_report_workbook.worksheets[0]

        for (cell_i,cell_j) in zip(edit_report_sheet['B'],edit_report_sheet['A']):
            i += 1
            if cell_i.value == _name_ and cell_j.value == _date_:
                edit_mode = 1
                try:
                    k
                except NameError:
                    k = i
                cells_A = 'A' + str(i + 1)
                cells_B = 'B' + str(i + 1)
                if not edit_report_sheet[cells_B].value == _name_ and edit_report_sheet[cells_A].value == _date_ or not edit_report_sheet[cells_A].value == _date_ or edit_report_sheet[cells_B].value == None:
                    edit_start_row = k
                    edit_row_count = i
                    edit_last_row = i

                    edit_row_count = edit_row_count - edit_start_row + 1
                    break
            elif cell_i.value == None:
                try:
                    edit_start_row = k
                except NameError:
                        messagebox.showwarning("確認","該当するデータは見つかりませんでした。")
                edit_row_count = i
                edit_row_count = edit_row_count - edit_start_row + 1
                break

        if "edit_start_row" in globals() and "edit_row_count" in globals():

            i = i - k
            i += 1
            if i >= 1:
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_1.current(m)
                d = confirmed_day_1 = f.day
                d = day_list.index(str(d))
                today_day_combobox_1.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_1 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_1.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_1 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_1.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（1行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_1 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_1.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h1 = confirmed_start_time_h_1 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_1.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m1 = confirmed_start_time_m_1 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_1.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h1 = confirmed_end_time_h_1 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_1.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m1 = confirmed_end_time_m_1 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_1.current(f)

            if i >= 2:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_2.current(m)
                d = confirmed_day_2 = f.day
                d = day_list.index(str(d))
                today_day_combobox_2.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_2 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_2.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_2 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_2.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（2行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_2 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_2.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h2 = confirmed_start_time_h_2 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_2.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m2 = confirmed_start_time_m_2 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_2.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h2 = confirmed_end_time_h_2 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_2.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m2 = confirmed_end_time_m_2 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_2.current(f)

            if i >= 3:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_3.current(m)
                d = confirmed_day_3 = f.day
                d = day_list.index(str(d))
                today_day_combobox_3.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_3 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_3.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_3 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_3.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（3行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_3 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_3.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h3 = confirmed_start_time_h_3 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_3.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m3 = confirmed_start_time_m_3 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_3.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h3 = confirmed_end_time_h_3 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_3.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m3 = confirmed_end_time_m_3 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_3.current(f)

            if i >= 4:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_4.current(m)
                d = confirmed_day_4 = f.day
                d = day_list.index(str(d))
                today_day_combobox_4.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_4 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_4.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_4 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_4.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（4行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_4 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_4.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h4 = confirmed_start_time_h_4 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_4.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m4 = confirmed_start_time_m_4 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_4.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h4 = confirmed_end_time_h_4 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_4.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m4 = confirmed_end_time_m_4 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_4.current(f)

            if i >= 5:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_5.current(m)
                d = confirmed_day_5 = f.day
                d = day_list.index(str(d))
                today_day_combobox_5.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_5 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_5.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_5 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_5.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（5行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_5 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_5.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h5 = confirmed_start_time_h_5 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_5.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m5 = confirmed_start_time_m_5 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_5.current(f)


                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h5 = confirmed_end_time_h_5 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_5.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m5 = confirmed_end_time_m_5 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_5.current(f)

            if i >= 6:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_6.current(m)
                d = confirmed_day_6 = f.day
                d = day_list.index(str(d))
                today_day_combobox_6.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_6 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_6.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_6 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_6.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（6行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_6 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_6.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h6 = confirmed_start_time_h_6 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_6.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m6 = confirmed_start_time_m_6 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_6.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h6 = confirmed_end_time_h_6 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_6.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m6 = confirmed_end_time_m_6 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_6.current(f)

            if i >= 7:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_7.current(m)
                d = confirmed_day_7 = f.day
                d = day_list.index(str(d))
                today_day_combobox_7.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_7 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_7.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_7 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_7.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（7行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_7 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_7.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h7 = confirmed_start_time_h_7 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_7.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m7 = confirmed_start_time_m_7 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_7.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h7 = confirmed_end_time_h_7 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_7.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m7 = confirmed_end_time_m_7 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_7.current(f)

            if i >= 8:
                k += 1
                cells_A = 'A' + str(k)
                f = edit_report_sheet[cells_A].value
                m = f.month
                m = month_list.index(str(m))
                today_month_combobox_8.current(m)
                d = confirmed_day_8 = f.day
                d = day_list.index(str(d))
                today_day_combobox_8.current(d)

                try:
                    cells_C = 'C' + str(k)
                    f =  confirmed_material_number_8 = edit_report_sheet[cells_C].value
                    f = material_numbers.index(str(f))
                    material_number_combobox_8.current(f)
                except ValueError:
                    pass
                try:
                    cells_D = 'D' + str(k)
                    f = confirmed_work_class_num_8 = edit_report_sheet[cells_D].value
                    f = work_class_number.index(str(f))
                    work_class_num_combobox_8.current(f)
                except ValueError:
                    messagebox.showwarning("確認","作業区分の詳細を入力していたため作業区分は削除されました、再度入力してください。（8行目）")

                cells_E = 'E' + str(k)
                f = confirmed_work_num_8 = edit_report_sheet[cells_E].value
                f = work_number.index(str(f))
                work_num_combobox_8.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_h8 = confirmed_start_time_h_8 = f.hour
                f = time_h_list.index(str(f))
                time_h_num_combobox_8.current(f)

                cells_F = 'F' + str(k)
                f = edit_report_sheet[cells_F].value
                f = start_time_m8 = confirmed_start_time_m_8 = f.minute
                f = time_m_list.index(str(f))
                time_m_num_combobox_8.current(f)


                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_h8 = confirmed_end_time_h_8 = f.hour
                f = time_h_list.index(str(f))
                endtime_h_num_combobox_8.current(f)

                cells_G = 'G' + str(k)
                f = edit_report_sheet[cells_G].value
                f = end_time_m8 = confirmed_end_time_m_8 = f.minute
                f = endtime_m_list.index(str(f))
                endtime_m_num_combobox_8.current(f)

        if 'end_time_m1' in globals():
            update_end_minute_1(k)
        if 'end_time_m2' in globals():
            update_end_minute_2(k)
        if 'end_time_m3' in globals():
            update_end_minute_3(k)
        if 'end_time_m4' in globals():
            update_end_minute_4(k)
        if 'end_time_m5' in globals():
            update_end_minute_5(k)
        if 'end_time_m6' in globals():
            update_end_minute_6(k)
        if 'end_time_m7' in globals():
            update_end_minute_7(k)
        if 'end_time_m8' in globals():
            update_end_minute_8(k)

        del k

def recalculate_all_rows():
    if 'end_time_m1' in globals():
        update_end_minute_1("k")
    if 'end_time_m2' in globals():
        update_end_minute_2("k")
    if 'end_time_m3' in globals():
        update_end_minute_3("k")
    if 'end_time_m4' in globals():
        update_end_minute_4("k")
    if 'end_time_m5' in globals():
        update_end_minute_5("k")
    if 'end_time_m6' in globals():
        update_end_minute_6("k")
    if 'end_time_m7' in globals():
        update_end_minute_7("k")
    if 'end_time_m8' in globals():
        update_end_minute_8("k")

# def overtime_check():
#     global overtime_check_mode
#     if (overtime_check_combobox['state'] == "noemal"):
#         overtime_check_combobox['state'] = "disabled"
#         overtime_check_mode = 0
#     else:
#         overtime_check_combobox['state'] = "noemal"
#         overtime_check_mode = 1

def set_full_day_entry():
    global confirmed_end_time_h_1
    global confirmed_end_time_m_1
    global end_time_h1
    global end_time_m1
    confirmed_end_time_h_1 = end_time_h1 = 17
    confirmed_end_time_m_1 = end_time_m1 = 30

    endtime_h_num_combobox_1.current(9)
    endtime_m_num_combobox_1.current(3)
    update_end_hour_1(0)
    update_end_minute_1(0)

def toggle_break_input_mode():
    global break_input_mode

    if break_listbox_1['state'] == "normal":
        break_listbox_1['state'] = "disabled"
        break_listbox_2['state'] = "disabled"
        break_listbox_3['state'] = "disabled"
        break_listbox_4['state'] = "disabled"
        break_listbox_5['state'] = "disabled"
        break_listbox_6['state'] = "disabled"
        break_listbox_7['state'] = "disabled"
        break_listbox_8['state'] = "disabled"
        break_input_mode = 0

        any_label.place_forget()
    else:
        any_label.place(x=620, y=8)
        break_listbox_1['state'] = "normal"
        break_listbox_2['state'] = "normal"
        break_listbox_3['state'] = "normal"
        break_listbox_4['state'] = "normal"
        break_listbox_5['state'] = "normal"
        break_listbox_6['state'] = "normal"
        break_listbox_7['state'] = "normal"
        break_listbox_8['state'] = "normal"
        break_input_mode = 1
        print(break_listbox_1['state'])
        pass

def clear_all_rows():

    global confirmed_name
    global break_input_mode
    global edit_mode
    global pending_rows
    if 'confirmed_name' in globals():
        del confirmed_name
    total_time_min_v.set("合計時間 : "+ "0" + " 分")
    overtime_time_min_v.set("残業時間 : "+ "0" + " 分")
    nightshift_minutes_label_text.set("夜勤時間 : "+ "0" + " 分")
    selected_name_label_text.set("名前選択")
    name_num_combobox.set("")
    # overtime_check_combobox.set("")
    any_label.place_forget()
    break_listbox_1['state'] = "disabled"
    break_listbox_2['state'] = "disabled"
    break_listbox_3['state'] = "disabled"
    break_listbox_4['state'] = "disabled"
    break_listbox_5['state'] = "disabled"
    break_listbox_6['state'] = "disabled"
    break_listbox_7['state'] = "disabled"
    break_listbox_8['state'] = "disabled"

    # overtime_check_combobox['state'] = "disabled"
    pending_rows = []
    overtime_check_mode = 0
    break_input_mode = 0
    nightshift_minutes = 0
    confirmed_overtime_check = 0
    edit_mode = 0

    clear_material_search()
    clear_row_1()
    clear_row_2()
    clear_row_3()
    clear_row_4()
    clear_row_5()
    clear_row_6()
    clear_row_7()
    clear_row_8()

def clear_row_1():
    global confirmed_material_number_1
    global confirmed_work_num_1
    global confirmed_work_class_num_1
    global confirmed_start_time_h_1
    global confirmed_start_time_m_1
    global confirmed_end_time_h_1
    global confirmed_end_time_m_1
    global confirmed_break_time_1
    global confirmed_nightshift_minutes_1
    global confirmed_work_class_detail_1
    global confirmed_month_1
    global confirmed_day_1
    global total_time_1
    global total_time
    global time_h_num_combobox_1
    global time_m_num_combobox_1
    global start_time_h1
    global start_time_m1
    global end_time_h1
    global end_time_m1
    global today_month_combobox_1
    global today_day_combobox_1

    if 'confirmed_material_number_1' in globals():
        del confirmed_material_number_1
    if 'confirmed_work_num_1' in globals():
        del confirmed_work_num_1
    if 'confirmed_start_time_h_1' in globals():
        start_time_h1 = confirmed_start_time_h_1 = 8
    if 'confirmed_start_time_m_1' in globals():
        if kanto_mode == 0:
            start_time_m1 = confirmed_start_time_m_1 = 10
        elif kanto_mode == 1:
            start_time_m1 = confirmed_start_time_m_1 = 00
    if 'confirmed_end_time_h_1' in globals():
        del confirmed_end_time_h_1
    if 'confirmed_end_time_m_1' in globals():
        del confirmed_end_time_m_1
    if 'total_time_1' in globals():
        total_time_1 = 0
    if 'total_time' in globals():
        total_time = 0

    if 'confirmed_break_time_1' in globals():
        confirmed_break_time_1 = 0
    if 'confirmed_work_class_num_1' in globals():
        confirmed_work_class_num_1 = 0
    if 'end_time_h1' in globals():
        del end_time_h1
    if 'end_time_m1' in globals():
        del end_time_m1
    if 'confirmed_work_class_detail_1' in globals():
        del confirmed_work_class_detail_1

    confirmed_nightshift_minutes_1 = 0
    today_month_combobox_1.current(month)
    today_day_combobox_1.current(day)
    if kanto_mode == 0:
        i = 0
        u = 1
    elif kanto_mode == 1:
        i = 0
        u = 0
    time_h_num_combobox_1.current(i)
    time_m_num_combobox_1.current(u)

    confirmed_month_1 = int(today_month_combobox_1.get())
    confirmed_day_1 = int(today_day_combobox_1.get())
    material_number_combobox_1.set("")
    work_num_combobox_1.set("")
    work_class_num_combobox_1.set("")
    endtime_h_num_combobox_1.set("")
    endtime_m_num_combobox_1.set("")
    total_label_v1.set("")
    break_listbox_1.set("")
    work_class_detail_entry_1.delete(0,tk.END)
    checkbutton_var_1.set("0")


def clear_row_2():
    global confirmed_material_number_2
    global confirmed_work_num_2
    global confirmed_start_time_h_2
    global confirmed_start_time_m_2
    global confirmed_end_time_h_2
    global confirmed_end_time_m_2
    global confirmed_nightshift_minutes_2
    global confirmed_work_class_detail_2
    global confirmed_month_2
    global confirmed_day_2
    global today_month_combobox_2
    global today_day_combobox_2
    global confirmed_break_time_2
    global end_time_h2
    global end_time_m2
    global total_time_2

    if 'confirmed_material_number_2' in globals():
        del confirmed_material_number_2
    if 'confirmed_work_num_2' in globals():
        del confirmed_work_num_2
    if 'confirmed_start_time_h_2' in globals():
        del confirmed_start_time_h_2
    if 'confirmed_start_time_m_2' in globals():
        del confirmed_start_time_m_2
    if 'confirmed_end_time_h_2' in globals():
        del confirmed_end_time_h_2
    if 'confirmed_end_time_m_2' in globals():
        del confirmed_end_time_m_2
    if 'total_time_2' in globals():
        total_time_2 = 0


    if 'confirmed_break_time_2' in globals():
        confirmed_break_time_2 = 0
    if 'confirmed_work_class_num_2' in globals():
        confirmed_work_class_num_2 = 0
    if 'end_time_h2' in globals():
        del end_time_h2
    if 'end_time_m2' in globals():
        del end_time_m2
    if 'confirmed_work_class_detail_2' in globals():
        del confirmed_work_class_detail_2

    confirmed_nightshift_minutes_2 = 0
    today_month_combobox_2.current(month)
    today_day_combobox_2.current(day)
    time_h_num_combobox_2.current(0)
    time_m_num_combobox_2.current(0)
    confirmed_month_2 = int(today_month_combobox_2.get())
    confirmed_day_2 = int(today_day_combobox_2.get())
    material_number_combobox_2.set("")
    work_num_combobox_2.set("")
    work_class_num_combobox_2.set("")
    time_h_num_combobox_2.set("")
    time_m_num_combobox_2.set("")
    endtime_h_num_combobox_2.set("")
    endtime_m_num_combobox_2.set("")
    break_listbox_2.set("")
    total_label_v2.set("")
    work_class_detail_entry_2.delete(0,tk.END)
    checkbutton_var_2.set("0")

def clear_row_3():
    global confirmed_material_number_3
    global confirmed_start_time_h_3
    global confirmed_work_num_3
    global confirmed_start_time_m_3
    global confirmed_end_time_h_3
    global confirmed_end_time_m_3
    global confirmed_nightshift_minutes_3
    global confirmed_work_class_detail_3
    global confirmed_month_3
    global confirmed_day_3
    global today_month_combobox_3
    global today_day_combobox_3
    global confirmed_break_time_3
    global end_time_h3
    global end_time_m3
    global total_time_3
    if 'confirmed_material_number_3' in globals():
        del confirmed_material_number_3
    if 'confirmed_work_num_3' in globals():
        del confirmed_work_num_3
    if 'confirmed_start_time_h_3' in globals():
        del confirmed_start_time_h_3
    if 'confirmed_start_time_m_3' in globals():
        del confirmed_start_time_m_3
    if 'confirmed_end_time_h_3' in globals():
        del confirmed_end_time_h_3
    if 'confirmed_end_time_m_3' in globals():
        del confirmed_end_time_m_3
    if 'total_time_3' in globals():
        total_time_3 = 0
    if 'confirmed_break_time_3' in globals():
        confirmed_break_time_3 = 0
    if 'confirmed_work_class_num_3' in globals():
        confirmed_work_class_num_3 = 0
    if 'end_time_h3' in globals():
        del end_time_h3
    if 'end_time_m3' in globals():
        del end_time_m3
    if 'confirmed_work_class_detail_3' in globals():
        del confirmed_work_class_detail_3

    confirmed_nightshift_minutes_3 = 0
    today_month_combobox_3.current(month)
    today_day_combobox_3.current(day)
    time_h_num_combobox_3.current(0)
    time_m_num_combobox_3.current(0)
    confirmed_month_3 = int(today_month_combobox_3.get())
    confirmed_day_3 = int(today_day_combobox_3.get())
    material_number_combobox_3.set("")
    work_num_combobox_3.set("")
    work_class_num_combobox_3.set("")
    time_h_num_combobox_3.set("")
    time_m_num_combobox_3.set("")
    endtime_h_num_combobox_3.set("")
    endtime_m_num_combobox_3.set("")
    break_listbox_3.set("")
    total_label_v3.set("")
    work_class_detail_entry_3.delete(0,tk.END)
    checkbutton_var_3.set("0")

def clear_row_4():
    global confirmed_material_number_4
    global confirmed_work_num_4
    global confirmed_start_time_h_4
    global confirmed_start_time_m_4
    global confirmed_end_time_h_4
    global confirmed_end_time_m_4
    global confirmed_nightshift_minutes_4
    global confirmed_work_class_detail_4
    global confirmed_month_4
    global confirmed_day_4
    global today_month_combobox_4
    global today_day_combobox_4
    global confirmed_break_time_4
    global end_time_h4
    global end_time_m4
    global total_time_4
    if 'confirmed_material_number_4' in globals():
        del confirmed_material_number_4
    if 'confirmed_work_num_4' in globals():
        del confirmed_work_num_4
    if 'confirmed_start_time_h_4' in globals():
        del confirmed_start_time_h_4
    if 'confirmed_start_time_m_4' in globals():
        del confirmed_start_time_m_4
    if 'confirmed_end_time_h_4' in globals():
        del confirmed_end_time_h_4
    if 'confirmed_end_time_m_4' in globals():
        del confirmed_end_time_m_4
    if 'total_time_4' in globals():
        total_time_4 = 0
    if 'confirmed_break_time_4' in globals():
        confirmed_break_time_4 = 0
    if 'confirmed_work_class_num_4' in globals():
        confirmed_work_class_num_4 = 0
    if 'end_time_h4' in globals():
        del end_time_h4
    if 'end_time_m4' in globals():
        del end_time_m4
    if 'confirmed_work_class_detail_4' in globals():
        del confirmed_work_class_detail_4

    confirmed_nightshift_minutes_4 = 0
    today_month_combobox_4.current(month)
    today_day_combobox_4.current(day)
    time_h_num_combobox_4.current(0)
    time_m_num_combobox_4.current(0)
    confirmed_month_4 = int(today_month_combobox_4.get())
    confirmed_day_4 = int(today_day_combobox_4.get())
    work_num_combobox_4.set("")
    work_class_num_combobox_4.set("")
    material_number_combobox_4.set("")
    time_h_num_combobox_4.set("")
    time_m_num_combobox_4.set("")
    endtime_h_num_combobox_4.set("")
    endtime_m_num_combobox_4.set("")
    break_listbox_4.set("")
    total_label_v4.set("")
    work_class_detail_entry_4.delete(0,tk.END)
    checkbutton_var_4.set("0")

def clear_row_5():
    global confirmed_material_number_5
    global confirmed_work_num_5
    global confirmed_start_time_h_5
    global confirmed_start_time_m_5
    global confirmed_end_time_h_5
    global confirmed_end_time_m_5
    global confirmed_nightshift_minutes_5
    global confirmed_work_class_detail_5
    global confirmed_month_5
    global confirmed_day_5
    global today_month_combobox_5
    global today_day_combobox_5
    global confirmed_break_time_5
    global end_time_h5
    global end_time_m5
    global total_time_5
    if 'confirmed_material_number_5' in globals():
        del confirmed_material_number_5
    if 'confirmed_work_num_5' in globals():
        del confirmed_work_num_5
    if 'confirmed_start_time_h_5' in globals():
        del confirmed_start_time_h_5
    if 'confirmed_start_time_m_5' in globals():
        del confirmed_start_time_m_5
    if 'confirmed_end_time_h_5' in globals():
        del confirmed_end_time_h_5
    if 'confirmed_end_time_m_5' in globals():
        del confirmed_end_time_m_5
    if 'total_time_5' in globals():
        total_time_5 = 0
    if 'confirmed_break_time_5' in globals():
        confirmed_break_time_5 = 0
    if 'confirmed_work_class_num_5' in globals():
        confirmed_work_class_num_5 = 0
    if 'end_time_h5' in globals():
        del end_time_h5
    if 'end_time_m5' in globals():
        del end_time_m5
    if 'confirmed_work_class_detail_5' in globals():
        del confirmed_work_class_detail_5

    confirmed_nightshift_minutes_5 = 0
    today_month_combobox_5.current(month)
    today_day_combobox_5.current(day)
    time_h_num_combobox_5.current(0)
    time_m_num_combobox_5.current(0)
    confirmed_month_5 = int(today_month_combobox_5.get())
    confirmed_day_5 = int(today_day_combobox_5.get())
    material_number_combobox_5.set("")
    work_num_combobox_5.set("")
    work_class_num_combobox_5.set("")
    time_h_num_combobox_5.set("")
    time_m_num_combobox_5.set("")
    endtime_h_num_combobox_5.set("")
    endtime_m_num_combobox_5.set("")
    break_listbox_5.set("")
    total_label_v5.set("")
    work_class_detail_entry_5.delete(0,tk.END)
    checkbutton_var_5.set("0")

def clear_row_6():
    global confirmed_material_number_6
    global confirmed_work_num_6
    global confirmed_start_time_h_6
    global confirmed_start_time_m_6
    global confirmed_end_time_h_6
    global confirmed_end_time_m_6
    global confirmed_nightshift_minutes_6
    global confirmed_work_class_detail_6
    global confirmed_month_6
    global confirmed_day_6
    global today_month_combobox_6
    global today_day_combobox_6
    global confirmed_break_time_6
    global end_time_h6
    global end_time_m6
    global total_time_6
    if 'confirmed_material_number_6' in globals():
        del confirmed_material_number_6
    if 'confirmed_work_num_6' in globals():
        del confirmed_work_num_6
    if 'confirmed_start_time_h_6' in globals():
        del confirmed_start_time_h_6
    if 'confirmed_start_time_m_6' in globals():
        del confirmed_start_time_m_6
    if 'confirmed_end_time_h_6' in globals():
        del confirmed_end_time_h_6
    if 'confirmed_end_time_m_6' in globals():
        del confirmed_end_time_m_6
    if 'total_time_6' in globals():
        total_time_6 = 0
    if 'confirmed_break_time_6' in globals():
        confirmed_break_time_6 = 0
    if 'confirmed_work_class_num_6' in globals():
        confirmed_work_class_num_6 = 0
    if 'end_time_h6' in globals():
        del end_time_h6
    if 'end_time_m6' in globals():
        del end_time_m6
    if 'confirmed_work_class_detail_6' in globals():
        del confirmed_work_class_detail_6

    confirmed_nightshift_minutes_6 = 0
    today_month_combobox_6.current(month)
    today_day_combobox_6.current(day)
    time_h_num_combobox_6.current(0)
    time_m_num_combobox_6.current(0)
    confirmed_month_6 = int(today_month_combobox_6.get())
    confirmed_day_6 = int(today_day_combobox_6.get())
    work_num_combobox_6.set("")
    work_class_num_combobox_6.set("")
    material_number_combobox_6.set("")
    time_h_num_combobox_6.set("")
    time_m_num_combobox_6.set("")
    endtime_h_num_combobox_6.set("")
    endtime_m_num_combobox_6.set("")
    break_listbox_6.set("")
    total_label_v6.set("")
    work_class_detail_entry_6.delete(0,tk.END)
    checkbutton_var_6.set("0")

def clear_row_7():
    global confirmed_material_number_7
    global confirmed_work_num_7
    global confirmed_start_time_h_7
    global confirmed_start_time_m_7
    global confirmed_end_time_h_7
    global confirmed_end_time_m_7
    global confirmed_nightshift_minutes_7
    global confirmed_work_class_detail_7
    global confirmed_month_7
    global confirmed_day_7
    global today_month_combobox_7
    global today_day_combobox_7
    global confirmed_break_time_7
    global end_time_h7
    global end_time_m7
    global total_time_7
    if 'confirmed_material_number_7' in globals():
        del confirmed_material_number_7
    if 'confirmed_work_num_7' in globals():
        del confirmed_work_num_7
    if 'confirmed_start_time_h_7' in globals():
        del confirmed_start_time_h_7
    if 'confirmed_start_time_m_7' in globals():
        del confirmed_start_time_m_7
    if 'confirmed_end_time_h_7' in globals():
        del confirmed_end_time_h_7
    if 'confirmed_end_time_m_7' in globals():
        del confirmed_end_time_m_7
    if 'total_time_7' in globals():
        total_time_7 = 0
    if 'confirmed_break_time_7' in globals():
        confirmed_break_time_7 = 0
    if 'confirmed_work_class_num_7' in globals():
        confirmed_work_class_num_7 = 0
    if 'end_time_h7' in globals():
        del end_time_h7
    if 'end_time_m7' in globals():
        del end_time_m7
    if 'confirmed_work_class_detail_7' in globals():
        del confirmed_work_class_detail_7

    confirmed_nightshift_minutes_7 = 0
    today_month_combobox_7.current(month)
    today_day_combobox_7.current(day)
    time_h_num_combobox_7.current(0)
    time_m_num_combobox_7.current(0)
    confirmed_month_7 = int(today_month_combobox_7.get())
    confirmed_day_7 = int(today_day_combobox_7.get())
    material_number_combobox_7.set("")
    work_num_combobox_7.set("")
    work_class_num_combobox_7.set("")
    time_h_num_combobox_7.set("")
    time_m_num_combobox_7.set("")
    endtime_h_num_combobox_7.set("")
    endtime_m_num_combobox_7.set("")
    break_listbox_7.set("")
    total_label_v7.set("")
    work_class_detail_entry_7.delete(0,tk.END)
    checkbutton_var_7.set("0")

def clear_row_8():
    global confirmed_material_number_8
    global confirmed_work_num_8
    global confirmed_start_time_h_8
    global confirmed_start_time_m_8
    global confirmed_end_time_h_8
    global confirmed_end_time_m_8
    global confirmed_nightshift_minutes_8
    global confirmed_work_class_detail_8
    global confirmed_month_8
    global confirmed_day_8
    global today_month_combobox_8
    global today_day_combobox_8
    global confirmed_break_time_8
    global end_time_h8
    global end_time_m8
    global total_time_8
    if 'confirmed_material_number_8' in globals():
        del confirmed_material_number_8
    if 'confirmed_work_num_8' in globals():
        del confirmed_work_num_8
    if 'confirmed_start_time_h_8' in globals():
        del confirmed_start_time_h_8
    if 'confirmed_start_time_m_8' in globals():
        del confirmed_start_time_m_8
    if 'confirmed_end_time_h_8' in globals():
        del confirmed_end_time_h_8
    if 'confirmed_end_time_m_8' in globals():
        del confirmed_end_time_m_8
    if 'total_time_8' in globals():
        total_time_8 = 0
    if 'confirmed_break_time_8' in globals():
        confirmed_break_time_8 = 0
    if 'confirmed_work_class_num_8' in globals():
        confirmed_work_class_num_8 = 0
    if 'end_time_h8' in globals():
        del end_time_h8
    if 'end_time_m8' in globals():
        del end_time_m8
    if 'confirmed_work_class_detail_8' in globals():
        del confirmed_work_class_detail_8

    confirmed_nightshift_minutes_8 = 0
    today_month_combobox_8.current(month)
    today_day_combobox_8.current(day)
    time_h_num_combobox_8.current(0)
    time_m_num_combobox_8.current(0)
    confirmed_month_8 = int(today_month_combobox_8.get())
    confirmed_day_8 = int(today_day_combobox_8.get())
    work_num_combobox_8.set("")
    work_class_num_combobox_8.set("")
    material_number_combobox_8.set("")
    time_h_num_combobox_8.set("")
    time_m_num_combobox_8.set("")
    endtime_h_num_combobox_8.set("")
    endtime_m_num_combobox_8.set("")
    break_listbox_8.set("")
    total_label_v8.set("")
    work_class_detail_entry_8.delete(0,tk.END)
    checkbutton_var_8.set("0")

def build_excel_row(today, material_number, work_class, work_number, start_hour, start_minute, end_hour, end_minute, total):
    start_time = datetime.time(start_hour,start_minute)
    end_time = datetime.time(end_hour,end_minute)
    confirmed = [today, confirmed_name, material_number, work_class, work_number, start_time, end_time, total, overtime_output_minutes, nightshift_output_minutes]
    return confirmed

def build_excel_row_without_material(today, work_class, work_number, start_hour, start_minute, end_hour, end_minute, total):
    start_time = datetime.time(start_hour,start_minute)
    end_time = datetime.time(end_hour,end_minute)
    confirmed = [today, confirmed_name,  "", work_class, work_number, start_time, end_time, total, overtime_output_minutes, nightshift_output_minutes]
    return confirmed

def register_entries():
    global target_row

    target_row = 0

    if 'confirmed_name' in globals():
        if'confirmed_start_time_h_1' in globals() and 'confirmed_start_time_m_1' in globals() and 'confirmed_end_time_h_1' in globals() and 'confirmed_end_time_m_1' in globals():
            report_workbook = load_required_workbook(report_workbook_path, "日報出力 Excel")
            report_sheet = report_workbook.worksheets[0]
            overtime_workbook = load_required_workbook(overtime_workbook_path, "残業確認 Excel")
            overtime_sheet = overtime_workbook.worksheets[0]
            if total_time >= 480:
                if overtime_check_mode == 1:
                    i = find_first_empty_row(overtime_sheet['A'], overtime_sheet)
                    list_i = [today, confirmed_name]
                    b = 1
                    for row_i in list_i:
                        overtime_sheet.cell(row=i,column=b).value = row_i
                        b+=1
                    try:
                        overtime_workbook.save(overtime_workbook_path)
                    except PermissionError:
                        messagebox.showwarning("確認","登録ファイルが開かれています。")
                        return

                target_row = find_first_empty_row(report_sheet['E'], report_sheet)
                i = 0
                for cell_i in report_sheet['E']:
                    i += 1
                    if target_row == i:
                        if cell_i.value == None:
                            break
                        else:
                            target_row +=1
                            break

                validate_row_1()

            else:
                MessageBox = messagebox.askyesno("確認","合計時間が480分を満たしていません\n登録してもよろしいでしょうか？")

                if MessageBox == True:
                    if overtime_check_mode == 1:
                        i = find_first_empty_row(overtime_sheet['A'], overtime_sheet)
                        list_i = [today, confirmed_name,]
                        b = 1
                        for row_i in list_i:
                            overtime_sheet.cell(row=i,column=b).value = row_i
                            b+=1
                        overtime_workbook.save(overtime_workbook_path)
                    target_row = find_first_empty_row(report_sheet['E'], report_sheet)
                    validate_row_1()

        else:
            messagebox.showwarning("確認","時間を選択してください")
    else:
        messagebox.showwarning("確認","名前を選択してください")

def validate_row_1():
    global edit_input_row_count
    if'confirmed_start_time_h_1' in globals() and 'confirmed_start_time_m_1' in globals() and 'confirmed_end_time_h_1' in globals() and 'confirmed_end_time_m_1' in globals():
        if edit_mode == 1:
            edit_input_row_count = 1
        if'confirmed_work_class_num_1' in globals() and 'confirmed_work_num_1' in globals():
            validate_row_2()
        else:
            messagebox.showwarning("確認","1行目 入力されていない箇所があります。")

def validate_row_2():
    global edit_input_row_count
    if'confirmed_start_time_h_2' in globals() and 'confirmed_start_time_m_2' in globals() and 'confirmed_end_time_h_2' in globals() and 'confirmed_end_time_m_2' in globals():
        if edit_mode == 1:
            edit_input_row_count = 2
        if 'confirmed_work_class_num_2' in globals() and 'confirmed_work_num_2' in globals():
            validate_row_3()
        else:
            messagebox.showwarning("確認","2行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def validate_row_3():
    global edit_input_row_count

    if'confirmed_start_time_h_3' in globals() and 'confirmed_start_time_m_3' in globals() and 'confirmed_end_time_h_3' in globals() and 'confirmed_end_time_m_3' in globals():
        if edit_mode == 1:
            edit_input_row_count = 3
        if 'confirmed_work_class_num_3' in globals() and 'confirmed_work_num_3' in globals():
            validate_row_4()
        else:
            messagebox.showwarning("確認","3行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def validate_row_4():
    global edit_input_row_count
    if'confirmed_start_time_h_4' in globals() and 'confirmed_start_time_m_4' in globals() and 'confirmed_end_time_h_4' in globals() and 'confirmed_end_time_m_4' in globals():
        if edit_mode == 1:
            edit_input_row_count = 4
        if 'confirmed_work_class_num_4' in globals() and 'confirmed_work_num_4' in globals():
            validate_row_5()
        else:
            messagebox.showwarning("確認","4行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def validate_row_5():
    global edit_input_row_count
    if'confirmed_start_time_h_5' in globals() and 'confirmed_start_time_m_5' in globals() and 'confirmed_end_time_h_5' in globals() and 'confirmed_end_time_m_5' in globals():
        if edit_mode == 1:
            edit_input_row_count = 5
        if 'confirmed_work_class_num_5' in globals() and 'confirmed_work_num_5' in globals():
            validate_row_6()
        else:
            messagebox.showwarning("確認","5行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def validate_row_6():
    global edit_input_row_count
    if'confirmed_start_time_h_6' in globals() and 'confirmed_start_time_m_6' in globals() and 'confirmed_end_time_h_6' in globals() and 'confirmed_end_time_m_6' in globals():
        if edit_mode == 1:
            edit_input_row_count = 6
        if 'confirmed_work_class_num_6' in globals() and 'confirmed_work_num_6' in globals():
            validate_row_7()
        else:
            messagebox.showwarning("確認","6行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def validate_row_7():
    global edit_input_row_count
    if'confirmed_start_time_h_7' in globals() and 'confirmed_start_time_m_7' in globals() and 'confirmed_end_time_h_7' in globals() and 'confirmed_end_time_m_7' in globals():
        if edit_mode == 1:
            edit_input_row_count = 7
        if 'confirmed_work_class_num_7' in globals() and 'confirmed_work_num_7' in globals():
            validate_row_8()
        else:
            messagebox.showwarning("確認","7行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def validate_row_8():
    global edit_input_row_count
    if'confirmed_start_time_h_8' in globals() and 'confirmed_start_time_m_8' in globals() and 'confirmed_end_time_h_8' in globals() and 'confirmed_end_time_m_8' in globals():
        if edit_mode == 1:

            edit_input_row_count = 8
        if 'confirmed_work_class_num_8' in globals() and 'confirmed_work_num_8' in globals():
            collect_row_1()
        else:
            messagebox.showwarning("確認","8行目 入力されていない箇所があります。")
    else:
        collect_row_1()

def collect_row_1():
    global pending_rows
    global overtime_output_minutes
    global nightshift_output_minutes
    global edit_input_row_count
    global edit_row_count

    All_contents_list = []

    # if edit_mode == 1:
    #     report_workbook = openpyxl.load_workbook(report_workbook_path, keep_vba=True)
    #     report_sheet = report_workbook.worksheets[0]

    #     target_row = edit_start_row

    if 'confirmed_end_time_m_2' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""
    if 'confirmed_end_time_m_2' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_1' in globals() and 'confirmed_work_class_num_1' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_1 + confirmed_work_class_detail_1
    elif 'confirmed_work_class_num_1' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_1 + ""

    if 'confirmed_start_time_h_1' in globals() and 'confirmed_start_time_m_1' in globals() and 'confirmed_end_time_h_1' in globals() and 'confirmed_end_time_m_1' in globals():
        if 'confirmed_material_number_1' in globals() and 'confirmed_work_num_1' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_1),
                                    int(confirmed_day_1)),
                                    confirmed_material_number_1,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_1,
                                    confirmed_start_time_h_1,
                                    confirmed_start_time_m_1,
                                    confirmed_end_time_h_1,
                                    confirmed_end_time_m_1,
                                    total_time_1
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_1),
                                    int(confirmed_day_1)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_1,
                                    confirmed_start_time_h_1,
                                    confirmed_start_time_m_1,
                                    confirmed_end_time_h_1,
                                    confirmed_end_time_m_1,
                                    total_time_1
                                    )
        pending_rows.append(All_contents_list)
        collect_row_2()

def collect_row_2():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes

    All_contents_list = []

    if 'confirmed_end_time_m_3' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""

    if 'confirmed_end_time_m_3' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_2' in globals() and 'confirmed_work_class_num_2' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_2 + confirmed_work_class_detail_2
    elif 'confirmed_work_class_num_2' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_2 + ""

    if 'confirmed_start_time_h_2' in globals() and 'confirmed_start_time_m_2' in globals() and 'confirmed_end_time_h_2' in globals() and 'confirmed_end_time_m_2' in globals():
        if 'confirmed_material_number_2' in globals() and 'confirmed_work_num_2' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_2),
                                    int(confirmed_day_2)),
                                    confirmed_material_number_2,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_2,
                                    confirmed_start_time_h_2,
                                    confirmed_start_time_m_2,
                                    confirmed_end_time_h_2,
                                    confirmed_end_time_m_2,
                                    total_time_2
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_2),
                                    int(confirmed_day_2)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_2,
                                    confirmed_start_time_h_2,
                                    confirmed_start_time_m_2,
                                    confirmed_end_time_h_2,
                                    confirmed_end_time_m_2,
                                    total_time_2
                                    )
        pending_rows.append(All_contents_list)
        collect_row_3()
    else:
        save_contents(pending_rows)


def collect_row_3():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes

    All_contents_list = []

    if 'confirmed_end_time_m_4' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""

    if 'confirmed_end_time_m_4' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_3' in globals() and 'confirmed_work_class_num_3' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_3 + confirmed_work_class_detail_3
    elif 'confirmed_work_class_num_3' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_3 + ""

    if 'confirmed_start_time_h_3' in globals() and 'confirmed_start_time_m_3' in globals() and 'confirmed_end_time_h_3' in globals() and 'confirmed_end_time_m_3' in globals():
        if 'confirmed_material_number_3' in globals() and 'confirmed_work_num_3' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_3),
                                    int(confirmed_day_3)),
                                    confirmed_material_number_3,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_3,
                                    confirmed_start_time_h_3,
                                    confirmed_start_time_m_3,
                                    confirmed_end_time_h_3,
                                    confirmed_end_time_m_3,
                                    total_time_3
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_3),
                                    int(confirmed_day_3)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_3,
                                    confirmed_start_time_h_3,
                                    confirmed_start_time_m_3,
                                    confirmed_end_time_h_3,
                                    confirmed_end_time_m_3,
                                    total_time_3
                                    )
        pending_rows.append(All_contents_list)
        collect_row_4()
    else:
        save_contents(pending_rows)

def collect_row_4():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes

    All_contents_list = []

    if 'confirmed_end_time_m_5' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""

    if 'confirmed_end_time_m_5' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_4' in globals() and 'confirmed_work_class_num_4' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_4 + confirmed_work_class_detail_4
    elif 'confirmed_work_class_num_4' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_4 + ""

    if 'confirmed_start_time_h_4' in globals() and 'confirmed_start_time_m_4' in globals() and 'confirmed_end_time_h_4' in globals() and 'confirmed_end_time_m_4' in globals():
        if 'confirmed_material_number_4' in globals() and 'confirmed_work_num_4' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_4),
                                    int(confirmed_day_4)),
                                    confirmed_material_number_4,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_4,
                                    confirmed_start_time_h_4,
                                    confirmed_start_time_m_4,
                                    confirmed_end_time_h_4,
                                    confirmed_end_time_m_4,
                                    total_time_4
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_4),
                                    int(confirmed_day_4)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_4,
                                    confirmed_start_time_h_4,
                                    confirmed_start_time_m_4,
                                    confirmed_end_time_h_4,
                                    confirmed_end_time_m_4,
                                    total_time_4
                                    )
        pending_rows.append(All_contents_list)
        # target_row+=1
        collect_row_5()
    else:
        messagebox.showwarning("確認","登録しました。")
        clear_all_rows()

def collect_row_5():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes
    All_contents_list = []
    if 'confirmed_end_time_m_6' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""

    if 'confirmed_end_time_m_6' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_5' in globals() and 'confirmed_work_class_num_5' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_5 + confirmed_work_class_detail_5
    elif 'confirmed_work_class_num_5' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_5 + ""

    if 'confirmed_start_time_h_5' in globals() and 'confirmed_start_time_m_5' in globals() and 'confirmed_end_time_h_5' in globals() and 'confirmed_end_time_m_5' in globals():
        if 'confirmed_material_number_5' in globals() and 'confirmed_work_num_5' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_5),
                                    int(confirmed_day_5)),
                                    confirmed_material_number_5,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_5,
                                    confirmed_start_time_h_5,
                                    confirmed_start_time_m_5,
                                    confirmed_end_time_h_5,
                                    confirmed_end_time_m_5,
                                    total_time_5
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_5),
                                    int(confirmed_day_5)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_5,
                                    confirmed_start_time_h_5,
                                    confirmed_start_time_m_5,
                                    confirmed_end_time_h_5,
                                    confirmed_end_time_m_5,
                                    total_time_5
                                    )
        pending_rows.append(All_contents_list)
        # target_row+=1
        collect_row_6()
    else:
        save_contents(pending_rows)

def collect_row_6():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes
    All_contents_list = []
    if 'confirmed_end_time_m_7' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""
    if 'confirmed_end_time_m_7' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_6' in globals() and 'confirmed_work_class_num_6' in globals():
        confirmed_work_class_detail  = confirmed_work_class_num_6 + confirmed_work_class_detail_6
    elif 'confirmed_work_class_num_6' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_6 + ""

    if 'confirmed_start_time_h_6' in globals() and 'confirmed_start_time_m_6' in globals() and 'confirmed_end_time_h_6' in globals() and 'confirmed_end_time_m_6' in globals():
        if 'confirmed_material_number_6' in globals() and 'confirmed_work_num_6' in globals():
           All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_6),
                                    int(confirmed_day_6)),
                                    confirmed_material_number_6,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_6,
                                    confirmed_start_time_h_6,
                                    confirmed_start_time_m_6,
                                    confirmed_end_time_h_6,
                                    confirmed_end_time_m_6,
                                    total_time_6
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_6),
                                    int(confirmed_day_6)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_6,
                                    confirmed_start_time_h_6,
                                    confirmed_start_time_m_6,
                                    confirmed_end_time_h_6,
                                    confirmed_end_time_m_6,
                                    total_time_6
                                    )
        pending_rows.append(All_contents_list)
        # target_row+=1
        collect_row_7()
    else:
        save_contents(pending_rows)

def collect_row_7():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes
    All_contents_list = []
    if 'confirmed_end_time_m_8' in globals():
        overtime_output_minutes = ""
    else:
        if overtime_time > 0:
            overtime_output_minutes = overtime_time
        else:
            overtime_output_minutes = ""

    if 'confirmed_end_time_m_8' in globals():
        nightshift_output_minutes = ""
    else:
        if confirmed_nightshift_minutes > 0:
            nightshift_output_minutes = confirmed_nightshift_minutes
        else:
            nightshift_output_minutes = ""

    if 'confirmed_work_class_detail_7' in globals() and 'confirmed_work_class_num_7' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_7 + confirmed_work_class_detail_7
    elif 'confirmed_work_class_num_7' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_7 + ""

    if 'confirmed_start_time_h_7' in globals() and 'confirmed_start_time_m_7' in globals() and 'confirmed_end_time_h_7' in globals() and 'confirmed_end_time_m_7' in globals():
        if 'confirmed_material_number_7' in globals() and 'confirmed_work_num_7' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_7),
                                    int(confirmed_day_7)),
                                    confirmed_material_number_7,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_7,
                                    confirmed_start_time_h_7,
                                    confirmed_start_time_m_7,
                                    confirmed_end_time_h_7,
                                    confirmed_end_time_m_7,
                                    total_time_7
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_7),
                                    int(confirmed_day_7)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_7,
                                    confirmed_start_time_h_7,
                                    confirmed_start_time_m_7,
                                    confirmed_end_time_h_7,
                                    confirmed_end_time_m_7,
                                    total_time_7
                                    )

        pending_rows.append(All_contents_list)
        # target_row+=1
        collect_row_8()
    else:
        save_contents(pending_rows)

def collect_row_8():
    global pending_rows
    global target_row
    global overtime_output_minutes
    global nightshift_output_minutes
    All_contents_list = []
    overtime_output_minutes = overtime_time
    nightshift_output_minutes = confirmed_nightshift_minutes

    if 'confirmed_work_class_detail_8' in globals() and 'confirmed_work_class_num_8' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_8 + confirmed_work_class_detail_8
    elif 'confirmed_work_class_num_8' in globals():
        confirmed_work_class_detail = confirmed_work_class_num_8 + ""

    if 'confirmed_start_time_h_8' in globals() and 'confirmed_start_time_m_8' in globals() and 'confirmed_end_time_h_8' in globals() and 'confirmed_end_time_m_8' in globals():
        if 'confirmed_material_number_8' in globals() and 'confirmed_work_num_8' in globals():
            All_contents_list = build_excel_row(datetime.date(current_year,
                                    int(confirmed_month_8),
                                    int(confirmed_day_8)),
                                    confirmed_material_number_8,
                                    confirmed_work_class_detail,
                                    confirmed_work_num_8,
                                    confirmed_start_time_h_8,
                                    confirmed_start_time_m_8,
                                    confirmed_end_time_h_8,
                                    confirmed_end_time_m_8,
                                    total_time_8
                                    )
        else:
            All_contents_list = build_excel_row_without_material(datetime.date(current_year,
                                    int(confirmed_month_8),
                                    int(confirmed_day_8)),
                                    confirmed_work_class_detail,
                                    confirmed_work_num_8,
                                    confirmed_start_time_h_8,
                                    confirmed_start_time_m_8,
                                    confirmed_end_time_h_8,
                                    confirmed_end_time_m_8,
                                    total_time_8
                                    )
    pending_rows.append(All_contents_list)
    save_contents(pending_rows)

def select_name(event):
    global confirmed_name
    i = name_number.index(name_num_combobox.get())
    confirmed_name = str(name_number[i])
    name_ward_str_h1 = str(name_number[i])
    selected_name_label_text.set(name_ward_str_h1)

def select_month_1(event):
    confirmed_month_1 = today_month_combobox_1.get()
    return confirmed_month_1

def select_month_2(event):
    confirmed_month_2 = today_month_combobox_2.get()
    return confirmed_month_2

def select_month_3(event):
    confirmed_month_3 = today_month_combobox_3.get()
    return confirmed_month_3

def select_month_4(event):
    confirmed_month_4 = today_month_combobox_4.get()
    return confirmed_month_4

def select_month_5(event):
    confirmed_month_5 = today_month_combobox_5.get()
    return confirmed_month_5

def select_month_6(event):
    confirmed_month_6 = today_month_combobox_6.get()
    return confirmed_month_6

def select_month_7(event):
    confirmed_month_7 = today_month_combobox_7.get()
    return confirmed_month_7

def select_month_8(event):
    confirmed_month_8 = today_month_combobox_8.get()
    return confirmed_month_8

def select_day_1(event):
    global confirmed_day_1
    confirmed_day_1 = int(today_day_combobox_1.get())

def select_day_2(event):
    global confirmed_day_2
    confirmed_day_2 = int(today_day_combobox_2.get())

def select_day_3(event):
    global confirmed_day_3
    confirmed_day_3 = int(today_day_combobox_3.get())

def select_day_4(event):
    global confirmed_day_4
    confirmed_day_4 = int(today_day_combobox_4.get())

def select_day_5(event):
    global confirmed_day_5
    confirmed_day_5 = int(today_day_combobox_5.get())

def select_day_6(event):
    global confirmed_day_6
    confirmed_day_6 = int(today_day_combobox_6.get())

def select_day_7(event):
    global confirmed_day_7
    confirmed_day_7 = int(today_day_combobox_7.get())

def select_day_8(event):
    global confirmed_day_8
    confirmed_day_8 = int(today_day_combobox_8.get())



def select_material_1(event):
    global confirmed_material_number_1
    confirmed_material_number_1 = material_number_combobox_1.get()

def select_material_2(event):
    global confirmed_material_number_2
    confirmed_material_number_2 = material_number_combobox_2.get()

def select_material_3(event):
    global confirmed_material_number_3
    confirmed_material_number_3 = material_number_combobox_3.get()

def select_material_4(event):
    global confirmed_material_number_4
    confirmed_material_number_4 = material_number_combobox_4.get()

def select_material_5(event):
    global confirmed_material_number_5
    confirmed_material_number_5 = material_number_combobox_5.get()

def select_material_6(event):
    global confirmed_material_number_6
    confirmed_material_number_6 = material_number_combobox_6.get()

def select_material_7(event):
    global confirmed_material_number_7
    confirmed_material_number_7 = material_number_combobox_7.get()

def select_material_8(event):
    global confirmed_material_number_8
    confirmed_material_number_8 = material_number_combobox_8.get()

def select_work_number_1(event):
    global confirmed_work_num_1
    confirmed_work_num_1 = work_num_combobox_1.get()

def select_work_number_2(event):
    global confirmed_work_num_2
    confirmed_work_num_2 = work_num_combobox_2.get()

def select_work_number_3(event):
    global confirmed_work_num_3
    confirmed_work_num_3 = work_num_combobox_3.get()

def select_work_number_4(event):
    global confirmed_work_num_4
    confirmed_work_num_4 = work_num_combobox_4.get()

def select_work_number_5(event):
    global confirmed_work_num_5
    confirmed_work_num_5 = work_num_combobox_5.get()

def select_work_number_6(event):
    global confirmed_work_num_6
    confirmed_work_num_6 = work_num_combobox_6.get()

def select_work_number_7(event):
    global confirmed_work_num_7
    confirmed_work_num_7 = work_num_combobox_7.get()

def select_work_number_8(event):
    global confirmed_work_num_8
    confirmed_work_num_8 = work_num_combobox_8.get()

def select_work_class_1(event):
    global confirmed_work_class_num_1
    confirmed_work_class_num_1 = work_class_num_combobox_1.get()

def select_work_class_2(event):
    global confirmed_work_class_num_2
    confirmed_work_class_num_2 = work_class_num_combobox_2.get()


def select_work_class_3(event):
    global confirmed_work_class_num_3
    confirmed_work_class_num_3 = work_class_num_combobox_3.get()

def select_work_class_4(event):
    global confirmed_work_class_num_4
    confirmed_work_class_num_4 = work_class_num_combobox_4.get()

def select_work_class_5(event):
    global confirmed_work_class_num_5
    confirmed_work_class_num_5 = work_class_num_combobox_5.get()

def select_work_class_6(event):
    global confirmed_work_class_num_6
    confirmed_work_class_num_6 = work_class_num_combobox_6.get()

def select_work_class_7(event):
    global confirmed_work_class_num_7
    confirmed_work_class_num_7 = work_class_num_combobox_7.get()

def select_work_class_8(event):
    global confirmed_work_class_num_8
    confirmed_work_class_num_8 = work_class_num_combobox_8.get()

def update_work_class_detail_1(event):
    global confirmed_work_class_detail_1
    confirmed_work_class_detail_1 = str(work_class_detail_entry_1.get())

def update_work_class_detail_2(event):
    global confirmed_work_class_detail_2
    confirmed_work_class_detail_2 = str(work_class_detail_entry_2.get())

def update_work_class_detail_3(event):
    global confirmed_work_class_detail_3
    confirmed_work_class_detail_3 = str(work_class_detail_entry_3.get())

def update_work_class_detail_4(event):
    global confirmed_work_class_detail_4
    confirmed_work_class_detail_4 = str(work_class_detail_entry_4.get())

def update_work_class_detail_5(event):
    global confirmed_work_class_detail_5
    confirmed_work_class_detail_5 = str(work_class_detail_entry_5.get())

def update_work_class_detail_6(event):
    global confirmed_work_class_detail_6
    confirmed_work_class_detail_6 = str(work_class_detail_entry_6.get())

def update_work_class_detail_7(event):
    global confirmed_work_class_detail_7
    confirmed_work_class_detail_7 = str(work_class_detail_entry_7.get())

def update_work_class_detail_8(event):
    global confirmed_work_class_detail_8
    confirmed_work_class_detail_8 = str(work_class_detail_entry_8.get())

def copy_material_to_row_1():
    global confirmed_material_number_2
    confirmed_material_number_2 = confirmed_material_number_1
    f = material_numbers.index(str(confirmed_material_number_1))
    material_number_combobox_2.current(f)

def copy_material_to_row_2():
    global confirmed_material_number_3
    confirmed_material_number_3 = confirmed_material_number_2
    f = material_numbers.index(str(confirmed_material_number_2))
    material_number_combobox_3.current(f)

def copy_material_to_row_3():
    global confirmed_material_number_4
    confirmed_material_number_4 = confirmed_material_number_3
    f = material_numbers.index(str(confirmed_material_number_3))
    material_number_combobox_4.current(f)

def copy_material_to_row_4():
    global confirmed_material_number_5
    confirmed_material_number_5 = confirmed_material_number_4
    f = material_numbers.index(str(confirmed_material_number_4))
    material_number_combobox_5.current(f)

def copy_material_to_row_5():
    global confirmed_material_number_6
    confirmed_material_number_6 = confirmed_material_number_5
    f = material_numbers.index(str(confirmed_material_number_5))
    material_number_combobox_6.current(f)

def copy_material_to_row_6():
    global confirmed_material_number_7
    confirmed_material_number_7 = confirmed_material_number_6
    f = material_numbers.index(str(confirmed_material_number_6))
    material_number_combobox_7.current(f)

def copy_material_to_row_7():
    global confirmed_material_number_8
    confirmed_material_number_8 = confirmed_material_number_7
    f = material_numbers.index(str(confirmed_material_number_7))
    material_number_combobox_8.current(f)

def update_start_hour_1(event):
    global start_time_h1
    global confirmed_start_time_h_1
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h1 = confirmed_start_time_h_1 = int(time_h_num_combobox_1.get())

    if 'start_time_h1' in globals() and 'start_time_m1' in globals() and 'end_time_h1' in globals() and 'end_time_m1' in globals():
        total_time_1,plus = calculate_work_minutes(start_time_h1, start_time_m1, end_time_h1, end_time_m1,int(confirmed_month_1),int(confirmed_day_1),int(confirmed_break_time_1))
        confirmed_nightshift_minutes_1 = calculate_nightshift_minutes(plus)
        total_label_v1.set(str(total_time_1) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_2(event):
    global start_time_h2
    global confirmed_start_time_h_2
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h2 = confirmed_start_time_h_2 = int(time_h_num_combobox_2.get())

    if 'start_time_h2' in globals() and 'start_time_m2' in globals() and 'end_time_h2' in globals() and 'end_time_m2' in globals():
        total_time_2,plus = calculate_work_minutes(start_time_h2, start_time_m2, end_time_h2, end_time_m2,int(confirmed_month_2),int(confirmed_day_2),int(confirmed_break_time_2))
        confirmed_nightshift_minutes_2 = calculate_nightshift_minutes(plus)

        total_label_v2.set(str(total_time_2) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_3(event):
    global start_time_h3
    global confirmed_start_time_h_3
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h3 = confirmed_start_time_h_3 = int(time_h_num_combobox_3.get())

    if 'start_time_h3' in globals() and 'start_time_m3' in globals() and 'end_time_h3' in globals() and 'end_time_m3' in globals():
        total_time_3,plus = calculate_work_minutes(start_time_h3, start_time_m3, end_time_h3, end_time_m3,int(confirmed_month_3),int(confirmed_day_3),int(confirmed_break_time_3))
        confirmed_nightshift_minutes_3 = calculate_nightshift_minutes(plus)

        total_label_v3.set(str(total_time_3) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_4(event):
    global start_time_h4
    global confirmed_start_time_h_4
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h4 = confirmed_start_time_h_4 = int(time_h_num_combobox_4.get())

    if 'start_time_h4' in globals() and 'start_time_m4' in globals() and 'end_time_h4' in globals() and 'end_time_m4' in globals():
        total_time_4,plus = calculate_work_minutes(start_time_h4, start_time_m4, end_time_h4, end_time_m4,int(confirmed_month_4),int(confirmed_day_4),int(confirmed_break_time_4))
        confirmed_nightshift_minutes_4 = calculate_nightshift_minutes(plus)

        total_label_v4.set(str(total_time_4) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_5(event):
    global start_time_h5
    global confirmed_start_time_h_5
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h5 = confirmed_start_time_h_5 = int(time_h_num_combobox_5.get())

    if 'start_time_h5' in globals() and 'start_time_m5' in globals() and 'end_time_h5' in globals() and 'end_time_m5' in globals():
        total_time_5,plus = calculate_work_minutes(start_time_h5, start_time_m5, end_time_h5, end_time_m5,int(confirmed_month_5),int(confirmed_day_5),int(confirmed_break_time_5))
        confirmed_nightshift_minutes_5 = calculate_nightshift_minutes(plus)

        total_label_v5.set(str(total_time_5) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_6(event):
    global start_time_h6
    global confirmed_start_time_h_6
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h6 = confirmed_start_time_h_6 = int(time_h_num_combobox_6.get())

    if 'start_time_h6' in globals() and 'start_time_m6' in globals() and 'end_time_h6' in globals() and 'end_time_m6' in globals():
        total_time_6,plus = calculate_work_minutes(start_time_h6, start_time_m6, end_time_h6, end_time_m6,int(confirmed_month_6),int(confirmed_day_6),int(confirmed_break_time_6))

        confirmed_nightshift_minutes_6 = calculate_nightshift_minutes(plus)

        total_label_v6.set(str(total_time_6) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_7(event):
    global start_time_h7
    global confirmed_start_time_h_7
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h7 = confirmed_start_time_h_7 = int(time_h_num_combobox_7.get())

    if 'start_time_h7' in globals() and 'start_time_m7' in globals() and 'end_time_h7' in globals() and 'end_time_m7' in globals():
        total_time_7,plus = calculate_work_minutes(start_time_h7, start_time_m7, end_time_h7, end_time_m7,int(confirmed_month_7),int(confirmed_day_7),int(confirmed_break_time_7))

        confirmed_nightshift_minutes_7 = calculate_nightshift_minutes(plus)

        total_label_v7.set(str(total_time_7) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_hour_8(event):
    global start_time_h8
    global confirmed_start_time_h_8
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_h8 = confirmed_start_time_h_8 = int(time_h_num_combobox_8.get())

    if 'start_time_h8' in globals() and 'start_time_m8' in globals() and 'end_time_h8' in globals() and 'end_time_m8' in globals():
        total_time_8,plus = calculate_work_minutes(start_time_h8, start_time_m8, end_time_h8, end_time_m8,int(confirmed_month_8),int(confirmed_day_8),int(confirmed_break_time_8))

        confirmed_nightshift_minutes_8 = calculate_nightshift_minutes(plus)

        total_label_v8.set(str(total_time_8) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_start_minute_1(event):
    global start_time_m1
    global confirmed_start_time_m_1
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m1 = confirmed_start_time_m_1 = int(time_m_num_combobox_1.get())

    if 'start_time_h1' in globals() and 'start_time_m1' in globals() and 'end_time_h1' in globals() and 'end_time_m1' in globals():
        total_time_1,plus = calculate_work_minutes(start_time_h1, start_time_m1, end_time_h1, end_time_m1,int(confirmed_month_1),int(confirmed_day_1),int(confirmed_break_time_1))
        confirmed_nightshift_minutes_1 = calculate_nightshift_minutes(plus)

        total_label_v1.set(str(total_time_1) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_2(event):
    global start_time_m2
    global confirmed_start_time_m_2
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m2 = confirmed_start_time_m_2 = int(time_m_num_combobox_2.get())

    if 'start_time_h2' in globals() and 'start_time_m2' in globals() and 'end_time_h2' in globals() and 'end_time_m2' in globals():
        total_time_2,plus = calculate_work_minutes(start_time_h2, start_time_m2, end_time_h2, end_time_m2,int(confirmed_month_2),int(confirmed_day_2),int(confirmed_break_time_2))
        confirmed_nightshift_minutes_2 = calculate_nightshift_minutes(plus)

        total_label_v2.set(str(total_time_2) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_3(event):
    global start_time_m3
    global confirmed_start_time_m_3
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m3 = confirmed_start_time_m_3 = int(time_m_num_combobox_3.get())

    if 'start_time_h3' in globals() and 'start_time_m3' in globals() and 'end_time_h3' in globals() and 'end_time_m3' in globals():
        total_time_3,plus = calculate_work_minutes(start_time_h3, start_time_m3, end_time_h3, end_time_m3,int(confirmed_month_3),int(confirmed_day_3),int(confirmed_break_time_3))
        confirmed_nightshift_minutes_3 = calculate_nightshift_minutes(plus)

        total_label_v3.set(str(total_time_3) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_4(event):
    global start_time_m4
    global confirmed_start_time_m_4
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m4 = confirmed_start_time_m_4 = int(time_m_num_combobox_4.get())

    if 'start_time_h4' in globals() and 'start_time_m4' in globals() and 'end_time_h4' in globals() and 'end_time_m4' in globals():
        total_time_4,plus = calculate_work_minutes(start_time_h4, start_time_m4, end_time_h4, end_time_m4,int(confirmed_month_4),int(confirmed_day_4),int(confirmed_break_time_4))
        confirmed_nightshift_minutes_4 = calculate_nightshift_minutes(plus)

        total_label_v4.set(str(total_time_4) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_5(event):
    global start_time_m5
    global confirmed_start_time_m_5
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m5 = confirmed_start_time_m_5 = int(time_m_num_combobox_5.get())

    if 'start_time_h5' in globals() and 'start_time_m5' in globals() and 'end_time_h5' in globals() and 'end_time_m5' in globals():
        total_time_5,plus = calculate_work_minutes(start_time_h5, start_time_m5, end_time_h5, end_time_m5,int(confirmed_month_5),int(confirmed_day_5),int(confirmed_break_time_5))
        confirmed_nightshift_minutes_5 = calculate_nightshift_minutes(plus)

        total_label_v5.set(str(total_time_5) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_6(event):
    global start_time_m6
    global confirmed_start_time_m_6
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m6 = confirmed_start_time_m_6 = int(time_m_num_combobox_6.get())

    if 'start_time_h6' in globals() and 'start_time_m6' in globals() and 'end_time_h6' in globals() and 'end_time_m6' in globals():
        total_time_6,plus = calculate_work_minutes(start_time_h6, start_time_m6, end_time_h6, end_time_m6,int(confirmed_month_6),int(confirmed_day_6),int(confirmed_break_time_6))

        confirmed_nightshift_minutes_6 = calculate_nightshift_minutes(plus)

        total_label_v6.set(str(total_time_6) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_7(event):
    global start_time_m7
    global confirmed_start_time_m_7
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m7 = confirmed_start_time_m_7 = int(time_m_num_combobox_7.get())

    if 'start_time_h7' in globals() and 'start_time_m7' in globals() and 'end_time_h7' in globals() and 'end_time_m7' in globals():
        total_time_7,plus = calculate_work_minutes(start_time_h7, start_time_m7, end_time_h7, end_time_m7,int(confirmed_month_7),int(confirmed_day_7),int(confirmed_break_time_7))

        confirmed_nightshift_minutes_7 = calculate_nightshift_minutes(plus)

        total_label_v7.set(str(total_time_7) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    # recalculate_all_rows()

def update_start_minute_8(event):
    global start_time_m8
    global confirmed_start_time_m_8
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    start_time_m8 = confirmed_start_time_m_8 = int(time_m_num_combobox_8.get())

    if 'start_time_h8' in globals() and 'start_time_m8' in globals() and 'end_time_h8' in globals() and 'end_time_m8' in globals():
        total_time_8,plus = calculate_work_minutes(start_time_h8, start_time_m8, end_time_h8, end_time_m8,int(confirmed_month_8),int(confirmed_day_8),int(confirmed_break_time_8))

        confirmed_nightshift_minutes_8 = calculate_nightshift_minutes(plus)

        total_label_v8.set(str(total_time_8) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_1(event):
    global end_time_h1
    global start_time_h2
    global end_time_h2
    global confirmed_end_time_h_1
    global confirmed_start_time_h_2
    global confirmed_end_time_h_2
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v2.set(endtime_h_num_combobox_1.get())
    endtime_h_num_v2.set(endtime_h_num_combobox_1.get())

    end_time_h1 = end_time_h2 = start_time_h2 = confirmed_end_time_h_1 = confirmed_start_time_h_2 = confirmed_end_time_h_2 = int(endtime_h_num_combobox_1.get())

    if 'start_time_h1' in globals() and 'start_time_m1' in globals() and 'end_time_h1' in globals() and 'end_time_m1' in globals():
        total_time_1,plus = calculate_work_minutes(start_time_h1, start_time_m1, end_time_h1, end_time_m1,int(confirmed_month_1),int(confirmed_day_1),int(confirmed_break_time_1))
        confirmed_nightshift_minutes_1 = calculate_nightshift_minutes(plus)

        total_label_v1.set(str(total_time_1) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_2(event):
    global end_time_h2
    global start_time_h3
    global end_time_h3
    global confirmed_end_time_h_2
    global confirmed_start_time_h_3
    global confirmed_end_time_h_3
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v3.set(endtime_h_num_combobox_2.get())
    endtime_h_num_v3.set(endtime_h_num_combobox_2.get())

    end_time_h2 = end_time_h3 = start_time_h3 = confirmed_end_time_h_2 = confirmed_start_time_h_3 = confirmed_end_time_h_3 = int(endtime_h_num_combobox_2.get())

    if 'start_time_h2' in globals() and 'start_time_m2' in globals() and 'end_time_h2' in globals() and 'end_time_m2' in globals():
        total_time_2,plus = calculate_work_minutes(start_time_h2, start_time_m2, end_time_h2, end_time_m2,int(confirmed_month_2),int(confirmed_day_2),int(confirmed_break_time_2))
        confirmed_nightshift_minutes_2 = calculate_nightshift_minutes(plus)

        total_label_v2.set(str(total_time_2) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_3(event):
    global end_time_h3
    global start_time_h4
    global end_time_h4
    global confirmed_end_time_h_3
    global confirmed_start_time_h_4
    global confirmed_end_time_h_4
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v4.set(endtime_h_num_combobox_3.get())
    endtime_h_num_v4.set(endtime_h_num_combobox_3.get())

    end_time_h3 = end_time_h4 = start_time_h4 = confirmed_end_time_h_3 = confirmed_start_time_h_4 = confirmed_end_time_h_4 = int(endtime_h_num_combobox_3.get())

    if 'start_time_h3' in globals() and 'start_time_m3' in globals() and 'end_time_h3' in globals() and 'end_time_m3' in globals():
        total_time_3,plus = calculate_work_minutes(start_time_h3, start_time_m3, end_time_h3, end_time_m3,int(confirmed_month_3),int(confirmed_day_3),int(confirmed_break_time_3))
        confirmed_nightshift_minutes_3 = calculate_nightshift_minutes(plus)

        total_label_v3.set(str(total_time_3) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_4(event):
    global end_time_h4
    global start_time_h5
    global end_time_h5
    global confirmed_end_time_h_4
    global confirmed_start_time_h_5
    global confirmed_end_time_h_5
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v5.set(endtime_h_num_combobox_4.get())
    endtime_h_num_v5.set(endtime_h_num_combobox_4.get())

    end_time_h4 = start_time_h5 = end_time_h5 = confirmed_end_time_h_4 = confirmed_start_time_h_5 = confirmed_end_time_h_5 = int(endtime_h_num_combobox_4.get())

    if 'start_time_h4' in globals() and 'start_time_m4' in globals() and 'end_time_h4' in globals() and 'end_time_m4' in globals():
        total_time_4,plus = calculate_work_minutes(start_time_h4, start_time_m4, end_time_h4, end_time_m4,int(confirmed_month_4),int(confirmed_day_4),int(confirmed_break_time_4))
        confirmed_nightshift_minutes_4 = calculate_nightshift_minutes(plus)

        total_label_v4.set(str(total_time_4) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_5(event):
    global end_time_h5
    global start_time_h6
    global end_time_h6
    global confirmed_end_time_h_5
    global confirmed_start_time_h_6
    global confirmed_end_time_h_6
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v6.set(endtime_h_num_combobox_5.get())
    endtime_h_num_v6.set(endtime_h_num_combobox_5.get())

    end_time_h5 = start_time_h6 = end_time_h6 = confirmed_end_time_h_5 = confirmed_start_time_h_6 = confirmed_end_time_h_6 = int(endtime_h_num_combobox_5.get())

    if 'start_time_h5' in globals() and 'start_time_m5' in globals() and 'end_time_h5' in globals() and 'end_time_m5' in globals():
        total_time_5,plus = calculate_work_minutes(start_time_h5, start_time_m5, end_time_h5, end_time_m5,int(confirmed_month_5),int(confirmed_day_5),int(confirmed_break_time_5))
        confirmed_nightshift_minutes_5 = calculate_nightshift_minutes(plus)

        total_label_v5.set(str(total_time_5) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_6(event):
    global end_time_h6
    global start_time_h7
    global end_time_h7
    global confirmed_end_time_h_6
    global confirmed_start_time_h_7
    global confirmed_end_time_h_7
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v7.set(endtime_h_num_combobox_6.get())
    endtime_h_num_v7.set(endtime_h_num_combobox_6.get())

    end_time_h6 = start_time_h7 = end_time_h7 = confirmed_end_time_h_6 = confirmed_start_time_h_7 = confirmed_end_time_h_7 = int(endtime_h_num_combobox_6.get())

    if 'start_time_h6' in globals() and 'start_time_m6' in globals() and 'end_time_h6' in globals() and 'end_time_m6' in globals():
        total_time_6,plus = calculate_work_minutes(start_time_h6, start_time_m6, end_time_h6, end_time_m6,int(confirmed_month_6),int(confirmed_day_6),int(confirmed_break_time_6))

        confirmed_nightshift_minutes_6 = calculate_nightshift_minutes(plus)

        total_label_v6.set(str(total_time_6) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_7(event):
    global end_time_h7
    global start_time_h8
    global end_time_h8
    global confirmed_end_time_h_7
    global confirmed_start_time_h_8
    global confirmed_end_time_h_8
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_v8.set(endtime_h_num_combobox_7.get())
    endtime_h_num_v8.set(endtime_h_num_combobox_7.get())

    end_time_h7 = start_time_h8 = end_time_h8 = confirmed_end_time_h_7 = confirmed_start_time_h_8 = confirmed_end_time_h_8 = int(endtime_h_num_combobox_7.get())

    if 'start_time_h7' in globals() and 'start_time_m7' in globals() and 'end_time_h7' in globals() and 'end_time_m7' in globals():
        total_time_7,plus = calculate_work_minutes(start_time_h7, start_time_m7, end_time_h7, end_time_m7,int(confirmed_month_7),int(confirmed_day_7),int(confirmed_break_time_7))

        confirmed_nightshift_minutes_7 = calculate_nightshift_minutes(plus)

        total_label_v7.set(str(total_time_7) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_hour_8(event):
    global end_time_h8
    global confirmed_end_time_h_8
    global confirmed_nightshift_minutes
    global overtime_time
    global total_time

    time_h_num_str_h8 = endtime_h_num_combobox_8.get()
    end_time_h8 = confirmed_end_time_h_8 = int(endtime_h_num_combobox_8.get())

    if 'start_time_h8' in globals() and 'start_time_m8' in globals() and 'end_time_h8' in globals() and 'end_time_m8' in globals():
        total_time_8,plus = calculate_work_minutes(start_time_h8, start_time_m8, end_time_h8, end_time_m8,int(confirmed_month_8),int(confirmed_day_8),int(confirmed_break_time_8))

        confirmed_nightshift_minutes_8 = calculate_nightshift_minutes(plus)

        total_label_v8.set(str(total_time_8) + " 分")
        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_end_minute_1(event):
    global end_time_m1
    global start_time_m2
    global total_time_1
    global confirmed_end_time_m_1
    global confirmed_start_time_m_2
    global confirmed_nightshift_minutes_1
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v2.set(endtime_m_num_combobox_1.get())
    end_time_m1 = start_time_m2 = confirmed_end_time_m_1 = confirmed_start_time_m_2 = int(endtime_m_num_combobox_1.get())

    if 'start_time_h1' in globals() and 'start_time_m1' in globals() and 'end_time_h1' in globals() and 'end_time_m1' in globals():
        total_time_1,plus = calculate_work_minutes(start_time_h1, start_time_m1, end_time_h1, end_time_m1,int(confirmed_month_1),int(confirmed_day_1),int(confirmed_break_time_1))
        confirmed_nightshift_minutes_1 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v1.set(str(total_time_1) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_2(event):
    global end_time_m2
    global start_time_m3
    global total_time_2
    global confirmed_end_time_m_2
    global confirmed_start_time_m_3
    global confirmed_nightshift_minutes_2
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v3.set(endtime_m_num_combobox_2.get())
    end_time_m2 = start_time_m3 = confirmed_end_time_m_2 = confirmed_start_time_m_3 = int(endtime_m_num_combobox_2.get())

    if 'start_time_h2' in globals() and 'start_time_m2' in globals() and 'end_time_h2' in globals() and 'end_time_m2' in globals():
        total_time_2,plus = calculate_work_minutes(start_time_h2, start_time_m2, end_time_h2, end_time_m2,int(confirmed_month_2),int(confirmed_day_2),int(confirmed_break_time_2))
        confirmed_nightshift_minutes_2 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v2.set(str(total_time_2) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_3(event):
    global end_time_m3
    global start_time_m4
    global total_time_3
    global confirmed_end_time_m_3
    global confirmed_start_time_m_4
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v4.set(endtime_m_num_combobox_3.get())
    end_time_m3 = start_time_m4 = confirmed_end_time_m_3 = confirmed_start_time_m_4 = int(endtime_m_num_combobox_3.get())

    if 'start_time_h3' in globals() and 'start_time_m3' in globals() and 'end_time_h3' in globals() and 'end_time_m3' in globals():
        total_time_3,plus = calculate_work_minutes(start_time_h3, start_time_m3, end_time_h3, end_time_m3,int(confirmed_month_3),int(confirmed_day_3),int(confirmed_break_time_3))
        confirmed_nightshift_minutes_3 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v3.set(str(total_time_3) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_4(event):
    global end_time_m4
    global start_time_m5
    global total_time_4
    global confirmed_end_time_m_4
    global confirmed_start_time_m_5
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v5.set(endtime_m_num_combobox_4.get())
    end_time_m4 = start_time_m5 = confirmed_end_time_m_4 = confirmed_start_time_m_5 = int(endtime_m_num_combobox_4.get())

    if 'start_time_h4' in globals() and 'start_time_m4' in globals() and 'end_time_h4' in globals() and 'end_time_m4' in globals():
        total_time_4,plus = calculate_work_minutes(start_time_h4, start_time_m4, end_time_h4, end_time_m4,int(confirmed_month_4),int(confirmed_day_4),int(confirmed_break_time_4))
        confirmed_nightshift_minutes_4 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v4.set(str(total_time_4) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_5(event):
    global end_time_m5
    global start_time_m6
    global total_time_5
    global confirmed_end_time_m_5
    global confirmed_start_time_m_6
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v6.set(endtime_m_num_combobox_5.get())
    end_time_m5 = start_time_m6 = confirmed_end_time_m_5 = confirmed_start_time_m_6 = int(endtime_m_num_combobox_5.get())

    if 'start_time_h5' in globals() and 'start_time_m5' in globals() and 'end_time_h5' in globals() and 'end_time_m5' in globals():
        total_time_5,plus = calculate_work_minutes(start_time_h5, start_time_m5, end_time_h5, end_time_m5,int(confirmed_month_5),int(confirmed_day_5),int(confirmed_break_time_5))
        confirmed_nightshift_minutes_5 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v5.set(str(total_time_5) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_6(event):
    global end_time_m6
    global start_time_m7
    global total_time_6
    global confirmed_end_time_m_6
    global confirmed_start_time_m_7
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v7.set(endtime_m_num_combobox_6.get())
    end_time_m6 = start_time_m7 = confirmed_end_time_m_6 = confirmed_start_time_m_7 = int(endtime_m_num_combobox_6.get())

    if 'start_time_h6' in globals() and 'start_time_m6' in globals() and 'end_time_h6' in globals() and 'end_time_m6' in globals():
        total_time_6,plus = calculate_work_minutes(start_time_h6, start_time_m6, end_time_h6, end_time_m6,int(confirmed_month_6),int(confirmed_day_6),int(confirmed_break_time_6))

        confirmed_nightshift_minutes_6 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v6.set(str(total_time_6) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_7(event):
    global end_time_m7
    global start_time_m8
    global total_time_7
    global confirmed_end_time_m_7
    global confirmed_start_time_m_8
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_v8.set(endtime_m_num_combobox_7.get())
    end_time_m7 = start_time_m8 = confirmed_end_time_m_7 = confirmed_start_time_m_8 = int(endtime_m_num_combobox_7.get())

    if 'start_time_h7' in globals() and 'start_time_m7' in globals() and 'end_time_h7' in globals() and 'end_time_m7' in globals():
        total_time_7,plus = calculate_work_minutes(start_time_h7, start_time_m7, end_time_h7, end_time_m7,int(confirmed_month_7),int(confirmed_day_7),int(confirmed_break_time_7))

        confirmed_nightshift_minutes_7 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v7.set(str(total_time_7) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_end_minute_8(event):
    global end_time_m8
    global total_time_8
    global confirmed_end_time_m_8
    global overtime_time
    global total_time
    global confirmed_nightshift_minutes

    time_m_num_str_m8 = endtime_m_num_combobox_8.get()
    end_time_m8  = confirmed_end_time_m_8 = int(endtime_m_num_combobox_8.get())

    if 'start_time_h8' in globals() and 'start_time_m8' in globals() and 'end_time_h8' in globals() and 'end_time_m8' in globals():
        total_time_8,plus = calculate_work_minutes(start_time_h8, start_time_m8, end_time_h8, end_time_m8,int(confirmed_month_8),int(confirmed_day_8),int(confirmed_break_time_8))

        confirmed_nightshift_minutes_8 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v8.set(str(total_time_8) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

def update_break_time_1(event):
    global confirmed_break_time_1
    global confirmed_nightshift_minutes
    global total_time_1
    global overtime_time
    global total_time

    confirmed_break_time_1 = break_listbox_1.get()

    if 'start_time_h1' in globals() and 'start_time_m1' in globals() and 'end_time_h1' in globals() and 'end_time_m1' in globals():
        total_time_1,plus = calculate_work_minutes(start_time_h1, start_time_m1, end_time_h1, end_time_m1,int(confirmed_month_1),int(confirmed_day_1),int(confirmed_break_time_1))
        confirmed_nightshift_minutes_1 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v1.set(str(total_time_1) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_break_time_2(event):
    global confirmed_break_time_2
    global confirmed_nightshift_minutes
    global total_time_2
    global overtime_time
    global total_time

    confirmed_break_time_2 = break_listbox_2.get()

    if 'start_time_h2' in globals() and 'start_time_m2' in globals() and 'end_time_h2' in globals() and 'end_time_m2' in globals():
        total_time_2,plus = calculate_work_minutes(start_time_h2, start_time_m2, end_time_h2, end_time_m2,int(confirmed_month_2),int(confirmed_day_2),int(confirmed_break_time_2))
        confirmed_nightshift_minutes_2 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v2.set(str(total_time_2) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_break_time_3(event):
    global confirmed_break_time_3
    global confirmed_nightshift_minutes
    global total_time_3
    global overtime_time
    global total_time

    confirmed_break_time_3 = break_listbox_3.get()

    if 'start_time_h3' in globals() and 'start_time_m3' in globals() and 'end_time_h3' in globals() and 'end_time_m3' in globals():
        total_time_3,plus = calculate_work_minutes(start_time_h3, start_time_m3, end_time_h3, end_time_m3,int(confirmed_month_3),int(confirmed_day_3),int(confirmed_break_time_3))
        confirmed_nightshift_minutes_3 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v3.set(str(total_time_3) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_break_time_4(event):
    global confirmed_break_time_4
    global confirmed_nightshift_minutes
    global total_time_4
    global overtime_time
    global total_time

    confirmed_break_time_4 = break_listbox_4.get()

    if 'start_time_h4' in globals() and 'start_time_m4' in globals() and 'end_time_h4' in globals() and 'end_time_m4' in globals():
        total_time_4,plus = calculate_work_minutes(start_time_h4, start_time_m4, end_time_h4, end_time_m4,int(confirmed_month_4),int(confirmed_day_4),int(confirmed_break_time_4))
        confirmed_nightshift_minutes_4 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v4.set(str(total_time_4) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_break_time_5(event):
    global confirmed_break_time_5
    global confirmed_nightshift_minutes
    global total_time_5
    global overtime_time
    global total_time

    confirmed_break_time_5 = break_listbox_5.get()

    if 'start_time_h5' in globals() and 'start_time_m5' in globals() and 'end_time_h5' in globals() and 'end_time_m5' in globals():
        total_time_5,plus = calculate_work_minutes(start_time_h5, start_time_m5, end_time_h5, end_time_m5,int(confirmed_month_5),int(confirmed_day_5),int(confirmed_break_time_5))
        confirmed_nightshift_minutes_5 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v5.set(str(total_time_5) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

    recalculate_all_rows()

def update_break_time_6(event):
    global confirmed_break_time_6
    global confirmed_nightshift_minutes
    global total_time_6
    global overtime_time
    global total_time

    confirmed_break_time_6 = break_listbox_6.get()

    if 'start_time_h6' in globals() and 'start_time_m6' in globals() and 'end_time_h6' in globals() and 'end_time_m6' in globals():
        total_time_6,plus = calculate_work_minutes(start_time_h6, start_time_m6, end_time_h6, end_time_m6,int(confirmed_month_6),int(confirmed_day_6),int(confirmed_break_time_6))
        confirmed_nightshift_minutes_6 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v6.set(str(total_time_6) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

        recalculate_all_rows()

def update_break_time_7(event):
    global confirmed_break_time_7
    global confirmed_nightshift_minutes
    global total_time_7
    global overtime_time
    global total_time

    confirmed_break_time_7 = break_listbox_7.get()

    if 'start_time_h7' in globals() and 'start_time_m7' in globals() and 'end_time_h7' in globals() and 'end_time_m7' in globals():
        total_time_7,plus = calculate_work_minutes(start_time_h7, start_time_m7, end_time_h7, end_time_m7,int(confirmed_month_7),int(confirmed_day_7),int(confirmed_break_time_7))

        confirmed_nightshift_minutes_7 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v7.set(str(total_time_7) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

        recalculate_all_rows()

def update_break_time_8(event):
    global confirmed_break_time_8
    global confirmed_nightshift_minutes
    global total_time_8
    global overtime_time
    global total_time

    confirmed_break_time_8 = break_listbox_8.get()

    if 'start_time_h8' in globals() and 'start_time_m8' in globals() and 'end_time_h8' in globals() and 'end_time_m8' in globals():
        total_time_8,plus = calculate_work_minutes(start_time_h8, start_time_m8, end_time_h8, end_time_m8,int(confirmed_month_8),int(confirmed_day_8),int(confirmed_break_time_8))

        confirmed_nightshift_minutes_8 = calculate_nightshift_minutes(plus)

        total_time = total_time_1 + total_time_2 + total_time_3 + total_time_4 + total_time_5 + total_time_6 + total_time_7 + total_time_8
        total_label_v8.set(str(total_time_8) + " 分")
        total_time_min_v.set("合計時間 : "+ (str(total_time)) + " 分")

        confirmed_nightshift_minutes = confirmed_nightshift_minutes_1 + confirmed_nightshift_minutes_2 + confirmed_nightshift_minutes_3 + confirmed_nightshift_minutes_4 + confirmed_nightshift_minutes_5 + confirmed_nightshift_minutes_6 + confirmed_nightshift_minutes_7 + confirmed_nightshift_minutes_8
        nightshift_minutes_label_text.set("夜勤時間 : "+ str(confirmed_nightshift_minutes) + " 分")

        overtime_time = update_overtime_minutes(total_time)

        recalculate_all_rows()

def calculate_work_minutes(start_hour,start_minute,end_hour,end_minute,work_month,work_day,break_minutes):
    global total_time
    global start_time
    global end_time
    global overtime_time
    global total_time
    if kanto_mode == 0:
        start_time = datetime.datetime(current_year,current_month,current_day,start_hour,start_minute)
        end_time = datetime.datetime(current_year,current_month,current_day,end_hour,end_minute)
        if end_time == time_23_59:
            plus = 1
        else:
            plus = 0
        delta = end_time - start_time
        total = (delta.total_seconds() / 60)
        if break_input_mode == 0:
            if break_10_10 > start_time:
                if break_15_10 <= end_time:
                    total -= 80
                elif break_13_00 <= end_time:
                    total -= 70
                elif break_10_10 <= end_time:
                    total -= 10
            elif break_13_00 > start_time:
                if break_15_10 <= end_time:
                    total -= 70
                elif break_13_00 <= end_time and break_15_10 >= end_time:
                    total -= 60
            elif break_15_10 > start_time:
                if break_15_10 < end_time:
                    total -= 10
        elif break_input_mode == 1:
            total -= break_minutes

        total += plus
        return total,plus


    elif kanto_mode == 1:
        start_time = datetime.datetime(current_year,current_month,current_day,start_hour,start_minute)
        end_time = datetime.datetime(current_year,current_month,current_day,end_hour,end_minute)
        if end_time == time_23_59:
            plus = 1
        else:
            plus = 0
        delta = end_time - start_time
        total = (delta.total_seconds() / 60)
        if break_input_mode == 0:
            if break_10_15 > start_time:
                if break_15_15 <= end_time:
                    total -= 90
                elif break_13_00 <= end_time:
                    total -= 75
                elif break_10_15 <= end_time:
                    total -= 15
            elif break_13_00 > start_time:
                if break_15_15 <= end_time:
                    total -= 75
                elif break_13_00 <= end_time and break_15_15 >= end_time:
                    total -= 60
            elif break_15_10 > start_time:
                if break_15_10 < end_time:
                    total -= 15
        elif break_input_mode == 1:
            total -= break_minutes

        total += plus
        return total,plus

def calculate_nightshift_minutes(plus):
    if  end_time > time_22_00:
        nightshift_minutes = end_time - time_22_00
        nightshift_minutes = (nightshift_minutes.total_seconds() / 60)
        nightshift_minutes += plus
    elif end_time <= time_5_00 and start_time <= time_5_00:
        nightshift_minutes = end_time - start_time
        nightshift_minutes = (nightshift_minutes.total_seconds() / 60)
    elif end_time > time_5_00 and start_time < time_5_00:
        nightshift_minutes = time_5_00 - start_time
        nightshift_minutes = (nightshift_minutes.total_seconds() / 60)
    else:
        nightshift_minutes = 0
    return nightshift_minutes

def update_overtime_minutes(total_time):
    if total_time > 480.0:
        overtime_time = total_time - 480
        overtime_time_min_v.set("残業時間 : "+ (str(overtime_time)) + " 分")
    else:
        overtime_time = 0
        overtime_time_min_v.set("残業時間 : "+ (str(overtime_time)) + " 分")
    return overtime_time


from openpyxl.styles.borders import Border, Side
from tkinter.ttk import Combobox
from tkinter import messagebox
from tkinter import *
import datetime
import openpyxl
import tkinter as tk
from zipfile import BadZipFile

def stop_startup(title, message):
    messagebox.showerror(title, message)
    raise SystemExit(1)

def get_required_workbook_path(path_cell, cell_address, workbook_role):
    path_value = path_cell.value
    if path_value is None or str(path_value).strip() == "":
        stop_startup(
            "起動エラー",
            "excelpath.xlsx の {0} に {1} のパスを設定してください。".format(
                cell_address,
                workbook_role,
            ),
        )
    return str(path_value).strip()

def load_required_workbook(workbook_path, workbook_role):
    try:
        return openpyxl.load_workbook(workbook_path, keep_vba=True)
    except FileNotFoundError:
        stop_startup(
            "起動エラー",
            "{0} が見つかりません。\nパス: {1}".format(workbook_role, workbook_path),
        )
    except PermissionError:
        stop_startup(
            "起動エラー",
            "{0} を開けません。\nExcel で開いている場合は閉じてから再実行してください。\nパス: {1}".format(
                workbook_role,
                workbook_path,
            ),
        )
    except (BadZipFile, OSError, openpyxl.utils.exceptions.InvalidFileException) as error:
        stop_startup(
            "起動エラー",
            "{0} を読み込めません。\nパス: {1}\n詳細: {2}".format(
                workbook_role,
                workbook_path,
                error,
            ),
        )

# 日付取得
dt_now = datetime.datetime.now()
current_year = dt_now.year
current_month = dt_now.month
current_day = dt_now.day
today = datetime.date(current_year,current_month,current_day)

# エクセルファイル読込

try:
    config_workbook = openpyxl.load_workbook('excelpath.xlsx', keep_vba=True)
except FileNotFoundError:
    stop_startup("起動エラー", "excelpath.xlsx が見つかりません。")
except PermissionError:
    stop_startup(
        "起動エラー",
        "excelpath.xlsx を開けません。\nExcel で開いている場合は閉じてから再実行してください。",
    )
except (BadZipFile, OSError, openpyxl.utils.exceptions.InvalidFileException) as error:
    stop_startup(
        "起動エラー",
        "excelpath.xlsx を読み込めません。\n詳細: {0}".format(error),
    )
config_sheet = config_workbook.worksheets[0]

master_path_cell = config_sheet['B1']

overtime_path_cell = config_sheet['B3']

kanto_value = config_sheet['D1'].value
class_value = config_sheet['D3'].value
name_value = config_sheet['D4'].value

if kanto_value == None:
    kanto_mode = 0
else:
    kanto_mode = int(kanto_value)

if class_value == None:
    class_mode = 0
else:
    class_mode = int(class_value)

if name_value == None:
    # name_value = 0
    pass
else:
    name_mode = int(name_value)

if class_mode == 0:
    work_name_column = "C"
    work_class_column = "N"
    work_number_column = "W"
    report_path_cell = config_sheet['B6']

if class_mode == 1:
    work_name_column = "D"
    work_class_column = "O"
    work_number_column = "X"
    report_path_cell = config_sheet['B7']

if class_mode == 2:
    work_name_column = "E"
    work_class_column = "P"
    work_number_column = "Y"
    report_path_cell = config_sheet['B8']

if class_mode == 3:
    work_name_column = "F"
    work_class_column = "Q"
    work_number_column = "Z"
    report_path_cell = config_sheet['B9']

if class_mode == 4:
    work_name_column = "G"
    work_class_column = "N"
    work_number_column = "W"
    report_path_cell = config_sheet['B10']

if class_mode == 5:
    work_name_column = "H"
    work_class_column = "P"
    work_number_column = "Y"
    report_path_cell = config_sheet['B11']

if class_mode == 6:
    work_name_column = "I"
    work_class_column = "R"
    work_number_column = "Z"
    report_path_cell = config_sheet['B12']

if class_mode == 7:
    work_name_column = "J"
    work_class_column = "S"
    work_number_column = "AA"
    report_path_cell = config_sheet['B13']

if class_mode == 8:
    work_name_column = "K"
    work_class_column = "T"
    work_number_column = "AB"
    report_path_cell = config_sheet['B14']

master_workbook_path = get_required_workbook_path(master_path_cell, "B1", "マスタ Excel")
report_workbook_path = get_required_workbook_path(report_path_cell, report_path_cell.coordinate, "日報出力 Excel")
overtime_workbook_path = get_required_workbook_path(overtime_path_cell, "B3", "残業確認 Excel")

master_workbook = load_required_workbook(master_workbook_path, "マスタ Excel")
side = Side(style='thin', color='000000')
border = Border(top=side, bottom=side, left=side, right=side)

master_sheet = master_workbook.worksheets[0]

break_10_10 = datetime.datetime(current_year,current_month,current_day,10,10)
break_10_15 = datetime.datetime(current_year,current_month,current_day,10,15)
break_13_00 = datetime.datetime(current_year,current_month,current_day,13,00)
break_15_10 = datetime.datetime(current_year,current_month,current_day,15,10)
break_15_15 = datetime.datetime(current_year,current_month,current_day,15,15)
time_22_00 = datetime.datetime(current_year,current_month,current_day,22,00)
time_23_59 = datetime.datetime(current_year,current_month,current_day,23,59)
time_5_00 = datetime.datetime(current_year,current_month,current_day,5,00)
time_0_00 = datetime.datetime(current_year,current_month,current_day,0,00)

total_time_1 = 0
total_time_2 = 0
total_time_3 = 0
total_time_4 = 0
total_time_5 = 0
total_time_6 = 0
total_time_7 = 0
total_time_8 = 0

confirmed_break_time_1 = 0
confirmed_break_time_2 = 0
confirmed_break_time_3 = 0
confirmed_break_time_4 = 0
confirmed_break_time_5 = 0
confirmed_break_time_6 = 0
confirmed_break_time_7 = 0
confirmed_break_time_8 = 0

confirmed_nightshift_minutes_1 = 0
confirmed_nightshift_minutes_2 = 0
confirmed_nightshift_minutes_3 = 0
confirmed_nightshift_minutes_4 = 0
confirmed_nightshift_minutes_5 = 0
confirmed_nightshift_minutes_6 = 0
confirmed_nightshift_minutes_7 = 0
confirmed_nightshift_minutes_8 = 0
# mode初期値
break_input_mode = 0
overtime_check_mode = 0
edit_mode = 0
pending_rows = []

# class_value = config_sheet['D1'].value

material_row = find_first_empty_row(master_sheet['A'], master_sheet)
material_row -= 1
material_numbers = []
material_number_labels = []
for row_1 in master_sheet["A2:A" + str(material_row)]:
    for col_1 in row_1:
        material_numbers.append(str(col_1.value)[:7])
        material_number_labels.append(str(col_1.value))
        material_search_values = [str(n) for n in material_number_labels]
        material_number_values = [str(n) for n in material_numbers]

work_class = find_first_empty_row(master_sheet[work_class_column], master_sheet)
work_class -= 1
work_class_number = []
for row_2 in master_sheet[work_class_column + "2:"+ work_class_column + str(work_class)]:
    for col_2 in row_2:
        work_class_number.append(str(col_2.value))

work = find_first_empty_row(master_sheet[work_number_column], master_sheet)
work -= 1
work_number = []
for row_3 in master_sheet[work_number_column + "2:" + work_number_column + str(work)]:
    for col_3 in row_3:
        work_number.append(col_3.value)
        work_number_str = [str(n) for n in work_number]

name_last_row = find_first_empty_row(master_sheet[work_name_column], master_sheet)
name_last_row -= 1
name_number = []
for row_5 in master_sheet[work_name_column + "2:"+ work_name_column + str(name_last_row)]:
    for col_5 in row_5:
        name_number.append(col_5.value)


# リスト作成
month_list = ["4","5","6","7","8","9","10","11","12","1","2","3"]
day_list = ["1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31"]
time_h_list = ["8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","0","1","2","3","4","5","6","7"]
time_m_list = ["0","10","20","30","40","50"]
time_m1_list = ["0","10","20","30","40","50"]
endtime_h_list = ["8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","0","1","2","3","4","5","6","7"]
endtime_m_list = ["0","10","20","30","40","50","59"]
break_list = ["0","10","20","30","40","50","60","70","80","90","100","110","120"]
overtime_check_list = ["0","30","60","90","120","180","210","270"]

# tkinter
version = "1.1.0"
fonts = "游ゴシック"

if kanto_mode == 0:
    app_title ='工数入力 ' + version
elif kanto_mode == 1:
    app_title ='関東工場 工数入力 ' + version

root1 = tk.Tk()
root1.title(app_title)
root1.geometry('1250x610')

name_frame = Frame(root1)
check_frame = Frame(root1,width=420, height=100, borderwidth=5, relief='ridge')
material_frame = Frame(root1)
tag_frame = Frame(root1)
work_frame = Frame(root1)
time_h_frame = Frame(root1)
time_m_frame = Frame(root1)
endtime_h_frame = Frame(root1)
endtime_m_frame = Frame(root1)
total_time_frame = LabelFrame(root1, text ='所要時間', width=90, height=345, borderwidth=5, relief='groove', font=(fonts,12,"bold"))
detail_frame = Frame(root1, width=100, height=40, borderwidth=5, relief='groove')
set_frame =  Frame(root1)
search_frame = Frame(root1, width=1190, height=120, borderwidth=5, relief='groove')

name_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=0, ipadx=700, ipady=40, sticky=NW)
tag_frame.grid(row=0, column=0, padx=0, pady=60, ipadx=500, ipady=0, sticky=NW)
material_frame.grid(row=0, column=0, padx=0, pady=90, ipadx=200, ipady=500, sticky=NW)
work_frame.grid(row=0, column=0, padx=610, pady=90, sticky=NW)
time_h_frame.grid(row=0, column=0, padx=780, pady=90, sticky=NW)
time_m_frame.grid(row=0, column=0, padx=840, pady=90, sticky=NW)
endtime_h_frame.grid(row=0, column=0, padx=920, pady=90, sticky=NW)
endtime_m_frame.grid(row=0, column=0, padx=980, pady=90, sticky=NW)
total_time_frame.grid(row=0, column=0, padx=1155, pady=67, sticky=NW)
total_time_frame.grid_propagate(0)
detail_frame.grid(row=0, column=0, padx=550, pady=415, ipadx=280, ipady=5, sticky=NW)
detail_frame.grid_propagate(0)
set_frame.grid(row=0, column=0, padx=800, pady=510, ipadx=0, sticky=NW)
search_frame.grid(row=0, column=0, padx=40, pady=410, ipadx=10, ipady=0, sticky=NW)

# --------------------------------------------------------------------------------
selected_name_label_text = StringVar()
selected_name_label_text.set("名前選択")
selected_name_label = tk.Label(name_frame, textvariable=selected_name_label_text, font=(fonts,20,"bold"))
selected_name_label.place(x=150, y=8)

name_num_v = StringVar()
name_num_combobox = Combobox(name_frame, textvariable=name_num_v, width=10, font=(fonts,15,"bold"))
name_num_combobox.grid(row=0, column=0, padx=0, pady=0, ipadx=0, sticky=W)

name_num_combobox.bind( '<<ComboboxSelected>>',select_name)

edit_button = tk.Button(name_frame, text="修正", command=load_rows_for_edit, font=(fonts,15,"bold"))
edit_button.place(x=1050, y=5)

allclearbutton = tk.Button(name_frame, text='削除', command=clear_all_rows, bg="red", font=(fonts,12,"bold"))
allclearbutton.grid(row=0, column=4, padx=180, pady=5, ipadx=30, ipady=5, sticky=W)

any_button = tk.Button(name_frame, text='休 憩 入 力',  bg="#ffc800", command=toggle_break_input_mode, font=(fonts,12,"bold"))
any_button.grid(row=0, column=3, padx=350, pady=5, ipadx=10, ipady=5, sticky=W)

any_label = tk.Label(name_frame, text='休憩時間入力モード',  bg="orange", font=(fonts,20,"bold"))

one_button = tk.Button(name_frame, text='1日一括入力',  bg="#ffc8ff", command=set_full_day_entry, font=(fonts,12,"bold"))
one_button.grid(row=0, column=3, padx=200, pady=5, ipadx=10, ipady=5, sticky=W)
# --------------------------------------------------------------------------------
tag_label_1 = tk.Label(tag_frame, text='日付', font=(fonts,15,"bold"))
tag_label_2 = tk.Label(tag_frame, text='工番', font=(fonts,15,"bold"))
tag_label_3 = tk.Label(tag_frame, text='作業区分', font=(fonts,15,"bold"))
tag_label_4 = tk.Label(tag_frame, text='詳細',foreground='red', font=(fonts,15,"bold"))
tag_label_5 = tk.Label(tag_frame, text='作業番号', font=(fonts,15,"bold"))
tag_label_6 = tk.Label(tag_frame, text='開始時間', font=(fonts,15,"bold"))
tag_label_7 = tk.Label(tag_frame, text='終了時間', font=(fonts,15,"bold"))
tag_label_8 = tk.Label(tag_frame, text='休憩', font=(fonts,15,"bold"))

# --------------------------------------------------------------------------------
today_month_combobox_1 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_2 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_3 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_4 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_5 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_6 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_7 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)
today_month_combobox_8 = Combobox(material_frame, width=2, font=(fonts,15,"bold"), height=10)

month_label_1 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_2 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_3 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_4 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_5 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_6 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_7 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))
month_label_8 = tk.Label(material_frame, text='月', font=(fonts,15,"bold"))

day_combobox_v_1 = StringVar()
day_combobox_v_2 = StringVar()
day_combobox_v_3 = StringVar()
day_combobox_v_4 = StringVar()
day_combobox_v_5 = StringVar()
day_combobox_v_6 = StringVar()
day_combobox_v_7 = StringVar()
day_combobox_v_8 = StringVar()

today_day_combobox_1 = Combobox(material_frame, textvariable = day_combobox_v_1, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_2 = Combobox(material_frame, textvariable = day_combobox_v_2, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_3 = Combobox(material_frame, textvariable = day_combobox_v_3, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_4 = Combobox(material_frame, textvariable = day_combobox_v_4, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_5 = Combobox(material_frame, textvariable = day_combobox_v_5, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_6 = Combobox(material_frame, textvariable = day_combobox_v_6, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_7 = Combobox(material_frame, textvariable = day_combobox_v_7, width=2, font=(fonts,15,"bold"), height=25)
today_day_combobox_8 = Combobox(material_frame, textvariable = day_combobox_v_8, width=2, font=(fonts,15,"bold"), height=25)

day_label_1 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_2 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_3 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_4 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_5 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_6 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_7 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))
day_label_8 = tk.Label(material_frame, text='日', font=(fonts,15,"bold"))

material_number_combobox_1 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_2 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_3 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_4 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_5 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_6 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_7 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)
material_number_combobox_8 = Combobox(material_frame, width=9, font=(fonts,15,"bold"), height=25)

checkbutton_var_1 = StringVar()
checkbutton_var_2 = StringVar()
checkbutton_var_3 = StringVar()
checkbutton_var_4 = StringVar()
checkbutton_var_5 = StringVar()
checkbutton_var_6 = StringVar()
checkbutton_var_7 = StringVar()
checkbutton_var_8 = StringVar()

checkbutton_var_1.set('0')
checkbutton_var_2.set('0')
checkbutton_var_3.set('0')
checkbutton_var_4.set('0')
checkbutton_var_5.set('0')
checkbutton_var_6.set('0')
checkbutton_var_7.set('0')
checkbutton_var_8.set('0')

layout_spacer_label = tk.Label(material_frame, "" ,font=("",0,"bold"))

material_search_checkbutton1 = tk.Checkbutton(material_frame, variable=checkbutton_var_1)
material_search_checkbutton2 = tk.Checkbutton(material_frame, variable=checkbutton_var_2)
material_search_checkbutton3 = tk.Checkbutton(material_frame, variable=checkbutton_var_3)
material_search_checkbutton4 = tk.Checkbutton(material_frame, variable=checkbutton_var_4)
material_search_checkbutton5 = tk.Checkbutton(material_frame, variable=checkbutton_var_5)
material_search_checkbutton6 = tk.Checkbutton(material_frame, variable=checkbutton_var_6)
material_search_checkbutton7 = tk.Checkbutton(material_frame, variable=checkbutton_var_7)
material_search_checkbutton8 = tk.Checkbutton(material_frame, variable=checkbutton_var_8)

row_copybutton_0 = tk.Button(material_frame, text='⤶', font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_1 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_1, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_2 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_2, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_3 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_3, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_4 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_4, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_5 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_5, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_6 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_6, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")
row_copybutton_7 = tk.Button(material_frame, text='⤶',command=copy_material_to_row_7, font=(fonts,12,"normal"), anchor=tk.CENTER, bg="#CCFFFF")

work_class_num_combobox_1 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_2 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_3 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_4 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_5 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_6 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_7 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)
work_class_num_combobox_8 = Combobox(material_frame, width=12, font=(fonts,15,"bold"), height=25)

work_class_detail_entry_1 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_2 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_3 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_4 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_5 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_6 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_7 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
work_class_detail_entry_8 = Entry(material_frame, width=10, font=(fonts,15,"bold"))
# --------------------------------------------------------------------------------
work_num_combobox_1 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_2 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_3 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_4 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_5 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_6 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_7 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
work_num_combobox_8 = Combobox(work_frame, width=12, font=(fonts,15,"bold"), height=25)
# --------------------------------------------------------------------------------
time_h_num_v1 = StringVar()
time_h_num_v2 = StringVar()
time_h_num_v3 = StringVar()
time_h_num_v4 = StringVar()
time_h_num_v5 = StringVar()
time_h_num_v6 = StringVar()
time_h_num_v7 = StringVar()
time_h_num_v8 = StringVar()

time_h_num_combobox_1 = Combobox(time_h_frame, textvariable=time_h_num_v1, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_2 = Combobox(time_h_frame, textvariable=time_h_num_v2, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_3 = Combobox(time_h_frame, textvariable=time_h_num_v3, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_4 = Combobox(time_h_frame, textvariable=time_h_num_v4, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_5 = Combobox(time_h_frame, textvariable=time_h_num_v5, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_6 = Combobox(time_h_frame, textvariable=time_h_num_v6, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_7 = Combobox(time_h_frame, textvariable=time_h_num_v7, width=3, font=(fonts,15,"bold"))
time_h_num_combobox_8 = Combobox(time_h_frame, textvariable=time_h_num_v8, width=3, font=(fonts,15,"bold"))
# --------------------------------------------------------------------------------
time_m_num_v1 = StringVar()
time_m_num_v2 = StringVar()
time_m_num_v3 = StringVar()
time_m_num_v4 = StringVar()
time_m_num_v5 = StringVar()
time_m_num_v6 = StringVar()
time_m_num_v7 = StringVar()
time_m_num_v8 = StringVar()
time_m_num_combobox_1 = Combobox(time_m_frame, textvariable=time_m_num_v1, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_2 = Combobox(time_m_frame, textvariable=time_m_num_v2, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_3 = Combobox(time_m_frame, textvariable=time_m_num_v3, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_4 = Combobox(time_m_frame, textvariable=time_m_num_v4, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_5 = Combobox(time_m_frame, textvariable=time_m_num_v5, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_6 = Combobox(time_m_frame, textvariable=time_m_num_v6, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_7 = Combobox(time_m_frame, textvariable=time_m_num_v7, width=3, font=(fonts,15,"bold"))
time_m_num_combobox_8 = Combobox(time_m_frame, textvariable=time_m_num_v8, width=3, font=(fonts,15,"bold"))
# --------------------------------------------------------------------------------
endtime_h_num_v1 = StringVar()
endtime_h_num_v2 = StringVar()
endtime_h_num_v3 = StringVar()
endtime_h_num_v4 = StringVar()
endtime_h_num_v5 = StringVar()
endtime_h_num_v6 = StringVar()
endtime_h_num_v7 = StringVar()
endtime_h_num_v8 = StringVar()
# --------------------------------------------------------------------------------
endtime_h_num_combobox_1 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v1, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_2 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v2, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_3 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v3, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_4 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v4, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_5 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v5, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_6 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v6, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_7 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v7, width=3, font=(fonts,15,"bold"))
endtime_h_num_combobox_8 = Combobox(endtime_h_frame, textvariable=endtime_h_num_v8, width=3, font=(fonts,15,"bold"))
# --------------------------------------------------------------------------------
endtime_m_num_v1 = StringVar()
endtime_m_num_v2 = StringVar()
endtime_m_num_v3 = StringVar()
endtime_m_num_v4 = StringVar()
endtime_m_num_v5 = StringVar()
endtime_m_num_v6 = StringVar()
endtime_m_num_v7 = StringVar()
endtime_m_num_v8 = StringVar()
# --------------------------------------------------------------------------------
endtime_m_num_combobox_1 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v1, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_2 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v2, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_3 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v3, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_4 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v4, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_5 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v5, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_6 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v6, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_7 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v7, width=3, font=(fonts,15,"bold"))
endtime_m_num_combobox_8 = Combobox(endtime_m_frame, textvariable=endtime_m_num_v8, width=3, font=(fonts,15,"bold"))

break_listbox_1 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_2 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_3 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_4 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_5 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_6 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_7 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))
break_listbox_8 = Combobox(endtime_m_frame, width=3,state="disabled", font=(fonts,15,"bold"))

row_clear_button_1 = tk.Button(endtime_m_frame, command=clear_row_1, bg="#FF4F02")
row_clear_button_2 = tk.Button(endtime_m_frame, command=clear_row_2, bg="#FF4F02")
row_clear_button_3 = tk.Button(endtime_m_frame, command=clear_row_3, bg="#FF4F02")
row_clear_button_4 = tk.Button(endtime_m_frame, command=clear_row_4, bg="#FF4F02")
row_clear_button_5 = tk.Button(endtime_m_frame, command=clear_row_5, bg="#FF4F02")
row_clear_button_6 = tk.Button(endtime_m_frame, command=clear_row_6, bg="#FF4F02")
row_clear_button_7 = tk.Button(endtime_m_frame, command=clear_row_7, bg="#FF4F02")
row_clear_button_8 = tk.Button(endtime_m_frame, command=clear_row_8, bg="#FF4F02")
# ----------------------------------------------------------------------------------------
total_label_v1 = StringVar()
total_label_v2 = StringVar()
total_label_v3 = StringVar()
total_label_v4 = StringVar()
total_label_v5 = StringVar()
total_label_v6 = StringVar()
total_label_v7 = StringVar()
total_label_v8 = StringVar()
# --------------------------------------------------------------------------------
total_label_1 = tk.Label(total_time_frame, textvariable=total_label_v1, font=(fonts,13,"bold"))
total_label_2 = tk.Label(total_time_frame, textvariable=total_label_v2, font=(fonts,13,"bold"))
total_label_3 = tk.Label(total_time_frame, textvariable=total_label_v3, font=(fonts,13,"bold"))
total_label_4 = tk.Label(total_time_frame, textvariable=total_label_v4, font=(fonts,13,"bold"))
total_label_5 = tk.Label(total_time_frame, textvariable=total_label_v5, font=(fonts,13,"bold"))
total_label_6 = tk.Label(total_time_frame, textvariable=total_label_v6, font=(fonts,13,"bold"))
total_label_7 = tk.Label(total_time_frame, textvariable=total_label_v7, font=(fonts,13,"bold"))
total_label_8 = tk.Label(total_time_frame, textvariable=total_label_v8, font=(fonts,13,"bold"))
# --------------------------------------------------------------------------------
total_time_min_v = StringVar()
total_time_min = tk.Label(detail_frame, textvariable=total_time_min_v, font=(fonts,15,"bold"))
total_time_min_v.set("合計時間 : "+ "0" + " 分")

overtime_time_min_v = StringVar()
overtime_time_min = tk.Label(detail_frame, textvariable=overtime_time_min_v, font=(fonts,15,"bold"))
overtime_time_min_v.set("残業時間 : "+ "0" + " 分")

nightshift_minutes_label_text = StringVar()
nightshift_minutes_label = tk.Label(detail_frame, textvariable=nightshift_minutes_label_text, font=(fonts,15,"bold"))
nightshift_minutes_label_text.set("夜勤時間 : "+ "0" + " 分")
# --------------------------------------------------------------------------------
set_button = tk.Button(set_frame, text='登　録', font=(fonts,20,"bold"), bg="#7FFF00", command=register_entries)
# --------------------------------------------------------------------------------
material_search_label = tk.Label(search_frame, text='工番検索', font=(fonts,15,"bold"))
material_search_label.place(x=1000, y=5)

material_search_entry = Entry(search_frame, width=30, font=(fonts,15,"bold"))
material_search_entry.grid(row=1, column=0, padx=10, pady=1)

material_search_button = tk.Button(search_frame, text='検　索', font=(fonts,12,"bold"), command=search_materials)
material_search_button.grid(row=1, column=1, padx=0, pady=1)

material_search_listbox = Listbox(search_frame,height=5, width=30, font=(fonts,15,"bold"))
material_search_listbox.grid(row=2, column=0, padx=14, pady=3)

material_search_button_clear = tk.Button(search_frame, text='挿　入', font=(fonts,12,"bold"), command=insert_selected_material)
material_search_button_clear.grid(row=2, column=1, padx=0, pady=1)
# # --------------------------------------------------------------------------------
name_num_combobox['values'] = name_number
try:
    name_num_combobox.current(name_mode)
    confirmed_name = name_num_combobox.get()
    selected_name_label_text.set(confirmed_name)
except NameError:
    pass
# --------------------------------------------------------------------------------
month = month_list.index(str(today.month))
today_month_combobox_1['values'] = month_list
today_month_combobox_1.current(month)
today_month_combobox_2['values'] = month_list
today_month_combobox_2.current(month)
today_month_combobox_3['values'] = month_list
today_month_combobox_3.current(month)
today_month_combobox_4['values'] = month_list
today_month_combobox_4.current(month)
today_month_combobox_5['values'] = month_list
today_month_combobox_5.current(month)
today_month_combobox_6['values'] = month_list
today_month_combobox_6.current(month)
today_month_combobox_7['values'] = month_list
today_month_combobox_7.current(month)
today_month_combobox_8['values'] = month_list
today_month_combobox_8.current(month)
confirmed_month_1 = int(today_month_combobox_1.get())
confirmed_month_2 = int(today_month_combobox_2.get())
confirmed_month_3 = int(today_month_combobox_3.get())
confirmed_month_4 = int(today_month_combobox_4.get())
confirmed_month_5 = int(today_month_combobox_5.get())
confirmed_month_6 = int(today_month_combobox_6.get())
confirmed_month_7 = int(today_month_combobox_7.get())
confirmed_month_8 = int(today_month_combobox_8.get())


day = day_list.index(str(today.day))
today_day_combobox_1['values'] = day_list
today_day_combobox_2['values'] = day_list
today_day_combobox_3['values'] = day_list
today_day_combobox_4['values'] = day_list
today_day_combobox_5['values'] = day_list
today_day_combobox_6['values'] = day_list
today_day_combobox_7['values'] = day_list
today_day_combobox_8['values'] = day_list

today_day_combobox_1.current(day)
today_day_combobox_2.current(day)
today_day_combobox_3.current(day)
today_day_combobox_4.current(day)
today_day_combobox_5.current(day)
today_day_combobox_6.current(day)
today_day_combobox_7.current(day)
today_day_combobox_8.current(day)

confirmed_day_1 = int(today_day_combobox_1.get())
confirmed_day_2 = int(today_day_combobox_2.get())
confirmed_day_3 = int(today_day_combobox_3.get())
confirmed_day_4 = int(today_day_combobox_4.get())
confirmed_day_5 = int(today_day_combobox_5.get())
confirmed_day_6 = int(today_day_combobox_6.get())
confirmed_day_7 = int(today_day_combobox_7.get())
confirmed_day_8 = int(today_day_combobox_8.get())

material_number_combobox_1['values'] = material_numbers
material_number_combobox_2['values'] = material_numbers
material_number_combobox_3['values'] = material_numbers
material_number_combobox_4['values'] = material_numbers
material_number_combobox_5['values'] = material_numbers
material_number_combobox_6['values'] = material_numbers
material_number_combobox_7['values'] = material_numbers
material_number_combobox_8['values'] = material_numbers

work_class_num_combobox_1['values'] = work_class_number
work_class_num_combobox_2['values'] = work_class_number
work_class_num_combobox_3['values'] = work_class_number
work_class_num_combobox_4['values'] = work_class_number
work_class_num_combobox_5['values'] = work_class_number
work_class_num_combobox_6['values'] = work_class_number
work_class_num_combobox_7['values'] = work_class_number
work_class_num_combobox_8['values'] = work_class_number

work_num_combobox_1['values'] = work_number
work_num_combobox_2['values'] = work_number
work_num_combobox_3['values'] = work_number
work_num_combobox_4['values'] = work_number
work_num_combobox_5['values'] = work_number
work_num_combobox_6['values'] = work_number
work_num_combobox_7['values'] = work_number
work_num_combobox_8['values'] = work_number

time_h_num_combobox_1['values'] = time_h_list
time_h_num_combobox_2['values'] = time_h_list
time_h_num_combobox_3['values'] = time_h_list
time_h_num_combobox_4['values'] = time_h_list
time_h_num_combobox_5['values'] = time_h_list
time_h_num_combobox_6['values'] = time_h_list
time_h_num_combobox_7['values'] = time_h_list
time_h_num_combobox_8['values'] = time_h_list

time_m_num_combobox_1['values'] = time_m1_list
time_m_num_combobox_2['values'] = time_m_list
time_m_num_combobox_3['values'] = time_m_list
time_m_num_combobox_4['values'] = time_m_list
time_m_num_combobox_5['values'] = time_m_list
time_m_num_combobox_6['values'] = time_m_list
time_m_num_combobox_7['values'] = time_m_list
time_m_num_combobox_8['values'] = time_m_list
if kanto_mode == 0:
    time_h_num_combobox_1.current(0)
    time_m_num_combobox_1.current(1)
elif kanto_mode == 1:
    time_h_num_combobox_1.current(0)
    time_m_num_combobox_1.current(0)
confirmed_start_time_h_1 = start_time_h1=int(time_h_num_combobox_1.get())
confirmed_start_time_m_1 = start_time_m1=int(time_m_num_combobox_1.get())

endtime_h_num_combobox_1['values'] = endtime_h_list
endtime_h_num_combobox_2['values'] = endtime_h_list
endtime_h_num_combobox_3['values'] = endtime_h_list
endtime_h_num_combobox_4['values'] = endtime_h_list
endtime_h_num_combobox_5['values'] = endtime_h_list
endtime_h_num_combobox_6['values'] = endtime_h_list
endtime_h_num_combobox_7['values'] = endtime_h_list
endtime_h_num_combobox_8['values'] = endtime_h_list

endtime_m_num_combobox_1['values'] = endtime_m_list
endtime_m_num_combobox_2['values'] = endtime_m_list
endtime_m_num_combobox_3['values'] = endtime_m_list
endtime_m_num_combobox_4['values'] = endtime_m_list
endtime_m_num_combobox_5['values'] = endtime_m_list
endtime_m_num_combobox_6['values'] = endtime_m_list
endtime_m_num_combobox_7['values'] = endtime_m_list
endtime_m_num_combobox_8['values'] = endtime_m_list

break_listbox_1['values'] = break_list
break_listbox_2['values'] = break_list
break_listbox_3['values'] = break_list
break_listbox_4['values'] = break_list
break_listbox_5['values'] = break_list
break_listbox_6['values'] = break_list
break_listbox_7['values'] = break_list
break_listbox_8['values'] = break_list

# overtime_check_combobox['values'] = overtime_check_list

# --------------------------------------------------------------------------------
name_num_combobox.bind( '<<ComboboxSelected>>')
# --------------------------------------------------------------------------------

today_month_combobox_1.bind( '<<ComboboxSelected>>',select_month_1)
today_month_combobox_2.bind( '<<ComboboxSelected>>',select_month_2)
today_month_combobox_3.bind( '<<ComboboxSelected>>',select_month_3)
today_month_combobox_4.bind( '<<ComboboxSelected>>',select_month_4)
today_month_combobox_5.bind( '<<ComboboxSelected>>',select_month_5)
today_month_combobox_6.bind( '<<ComboboxSelected>>',select_month_6)
today_month_combobox_7.bind( '<<ComboboxSelected>>',select_month_7)
today_month_combobox_8.bind( '<<ComboboxSelected>>',select_month_8)

today_day_combobox_1.bind( '<<ComboboxSelected>>',select_day_1)
today_day_combobox_2.bind( '<<ComboboxSelected>>',select_day_2)
today_day_combobox_3.bind( '<<ComboboxSelected>>',select_day_3)
today_day_combobox_4.bind( '<<ComboboxSelected>>',select_day_4)
today_day_combobox_5.bind( '<<ComboboxSelected>>',select_day_5)
today_day_combobox_6.bind( '<<ComboboxSelected>>',select_day_6)
today_day_combobox_7.bind( '<<ComboboxSelected>>',select_day_7)
today_day_combobox_8.bind( '<<ComboboxSelected>>',select_day_8)

material_number_combobox_1.bind( '<<ComboboxSelected>>', select_material_1)
material_number_combobox_2.bind( '<<ComboboxSelected>>', select_material_2)
material_number_combobox_3.bind( '<<ComboboxSelected>>', select_material_3)
material_number_combobox_4.bind( '<<ComboboxSelected>>', select_material_4)
material_number_combobox_5.bind( '<<ComboboxSelected>>', select_material_5)
material_number_combobox_6.bind( '<<ComboboxSelected>>', select_material_6)
material_number_combobox_7.bind( '<<ComboboxSelected>>', select_material_7)
material_number_combobox_8.bind( '<<ComboboxSelected>>', select_material_8)

work_num_combobox_1.bind( '<<ComboboxSelected>>', select_work_number_1)
work_num_combobox_2.bind( '<<ComboboxSelected>>', select_work_number_2)
work_num_combobox_3.bind( '<<ComboboxSelected>>', select_work_number_3)
work_num_combobox_4.bind( '<<ComboboxSelected>>', select_work_number_4)
work_num_combobox_5.bind( '<<ComboboxSelected>>', select_work_number_5)
work_num_combobox_6.bind( '<<ComboboxSelected>>', select_work_number_6)
work_num_combobox_7.bind( '<<ComboboxSelected>>', select_work_number_7)
work_num_combobox_8.bind( '<<ComboboxSelected>>', select_work_number_8)

work_class_num_combobox_1.bind( '<<ComboboxSelected>>', select_work_class_1)
work_class_num_combobox_2.bind( '<<ComboboxSelected>>', select_work_class_2)
work_class_num_combobox_3.bind( '<<ComboboxSelected>>', select_work_class_3)
work_class_num_combobox_4.bind( '<<ComboboxSelected>>', select_work_class_4)
work_class_num_combobox_5.bind( '<<ComboboxSelected>>', select_work_class_5)
work_class_num_combobox_6.bind( '<<ComboboxSelected>>', select_work_class_6)
work_class_num_combobox_7.bind( '<<ComboboxSelected>>', select_work_class_7)
work_class_num_combobox_8.bind( '<<ComboboxSelected>>', select_work_class_8)

work_class_detail_entry_1.bind( '<Motion>', update_work_class_detail_1)
work_class_detail_entry_2.bind( '<Motion>', update_work_class_detail_2)
work_class_detail_entry_3.bind( '<Motion>', update_work_class_detail_3)
work_class_detail_entry_4.bind( '<Motion>', update_work_class_detail_4)
work_class_detail_entry_5.bind( '<Motion>', update_work_class_detail_5)
work_class_detail_entry_6.bind( '<Motion>', update_work_class_detail_6)
work_class_detail_entry_7.bind( '<Motion>', update_work_class_detail_7)
work_class_detail_entry_8.bind( '<Motion>', update_work_class_detail_8)

time_h_num_combobox_1.bind( '<<ComboboxSelected>>', update_start_hour_1)
time_h_num_combobox_2.bind( '<<ComboboxSelected>>', update_start_hour_2)
time_h_num_combobox_3.bind( '<<ComboboxSelected>>', update_start_hour_3)
time_h_num_combobox_4.bind( '<<ComboboxSelected>>', update_start_hour_4)
time_h_num_combobox_5.bind( '<<ComboboxSelected>>', update_start_hour_5)
time_h_num_combobox_6.bind( '<<ComboboxSelected>>', update_start_hour_6)
time_h_num_combobox_7.bind( '<<ComboboxSelected>>', update_start_hour_7)
time_h_num_combobox_8.bind( '<<ComboboxSelected>>', update_start_hour_8)

time_m_num_combobox_1.bind( '<<ComboboxSelected>>', update_start_minute_1)
time_m_num_combobox_2.bind( '<<ComboboxSelected>>', update_start_minute_2)
time_m_num_combobox_3.bind( '<<ComboboxSelected>>', update_start_minute_3)
time_m_num_combobox_4.bind( '<<ComboboxSelected>>', update_start_minute_4)
time_m_num_combobox_5.bind( '<<ComboboxSelected>>', update_start_minute_5)
time_m_num_combobox_6.bind( '<<ComboboxSelected>>', update_start_minute_6)
time_m_num_combobox_7.bind( '<<ComboboxSelected>>', update_start_minute_7)
time_m_num_combobox_8.bind( '<<ComboboxSelected>>', update_start_minute_8)


endtime_h_num_combobox_1.bind( '<<ComboboxSelected>>', update_end_hour_1)
endtime_h_num_combobox_2.bind( '<<ComboboxSelected>>', update_end_hour_2)
endtime_h_num_combobox_3.bind( '<<ComboboxSelected>>', update_end_hour_3)
endtime_h_num_combobox_4.bind( '<<ComboboxSelected>>', update_end_hour_4)
endtime_h_num_combobox_5.bind( '<<ComboboxSelected>>', update_end_hour_5)
endtime_h_num_combobox_6.bind( '<<ComboboxSelected>>', update_end_hour_6)
endtime_h_num_combobox_7.bind( '<<ComboboxSelected>>', update_end_hour_7)
endtime_h_num_combobox_8.bind( '<<ComboboxSelected>>', update_end_hour_8)


endtime_m_num_combobox_1.bind( '<<ComboboxSelected>>', update_end_minute_1)
endtime_m_num_combobox_2.bind( '<<ComboboxSelected>>', update_end_minute_2)
endtime_m_num_combobox_3.bind( '<<ComboboxSelected>>', update_end_minute_3)
endtime_m_num_combobox_4.bind( '<<ComboboxSelected>>', update_end_minute_4)
endtime_m_num_combobox_5.bind( '<<ComboboxSelected>>', update_end_minute_5)
endtime_m_num_combobox_6.bind( '<<ComboboxSelected>>', update_end_minute_6)
endtime_m_num_combobox_7.bind( '<<ComboboxSelected>>', update_end_minute_7)
endtime_m_num_combobox_8.bind( '<<ComboboxSelected>>', update_end_minute_8)


break_listbox_1.bind( '<<ComboboxSelected>>', update_break_time_1)
break_listbox_2.bind( '<<ComboboxSelected>>', update_break_time_2)
break_listbox_3.bind( '<<ComboboxSelected>>', update_break_time_3)
break_listbox_4.bind( '<<ComboboxSelected>>', update_break_time_4)
break_listbox_5.bind( '<<ComboboxSelected>>', update_break_time_5)
break_listbox_6.bind( '<<ComboboxSelected>>', update_break_time_6)
break_listbox_7.bind( '<<ComboboxSelected>>', update_break_time_7)
break_listbox_8.bind( '<<ComboboxSelected>>', update_break_time_8)


# --------------------------------------------------------------------------------

# --------------------------------------------------------------------------------
tag_label_1.grid(row=0, column=0, padx=50)
tag_label_2.place(x=205, y=1)
tag_label_3.place(x=364, y=1)
tag_label_4.place(x=533, y=1)
tag_label_5.place(x=655, y=1)
tag_label_6.place(x=791, y=1)
tag_label_7.place(x=933, y=1)
tag_label_8.place(x=1040, y=1)

today_month_combobox_1.grid(row=0, column=0, padx=10, pady=4)
today_month_combobox_2.grid(row=1, column=0, padx=10, pady=4)
today_month_combobox_3.grid(row=2, column=0, padx=10, pady=4)
today_month_combobox_4.grid(row=3, column=0, padx=10, pady=4)
today_month_combobox_5.grid(row=4, column=0, padx=10, pady=4)
today_month_combobox_6.grid(row=5, column=0, padx=10, pady=4)
today_month_combobox_7.grid(row=6, column=0, padx=10, pady=4)
today_month_combobox_8.grid(row=7, column=0, padx=10, pady=4)

month_label_1.place(x=54.5, y=5)
month_label_2.place(x=54.5, y=45)
month_label_3.place(x=54.5, y=85)
month_label_4.place(x=54.5, y=125)
month_label_5.place(x=54.5, y=165)
month_label_6.place(x=54.5, y=205)
month_label_7.place(x=54.5, y=245)
month_label_8.place(x=54.5, y=285)

today_day_combobox_1.grid(row=0, column=2, padx=25, pady=4)
today_day_combobox_2.grid(row=1, column=2, padx=25, pady=4)
today_day_combobox_3.grid(row=2, column=2, padx=25, pady=4)
today_day_combobox_4.grid(row=3, column=2, padx=25, pady=4)
today_day_combobox_5.grid(row=4, column=2, padx=25, pady=4)
today_day_combobox_6.grid(row=5, column=2, padx=25, pady=4)
today_day_combobox_7.grid(row=6, column=2, padx=25, pady=4)
today_day_combobox_8.grid(row=7, column=2, padx=25, pady=4)

day_label_1.place(x=135, y=5)
day_label_2.place(x=135, y=45)
day_label_3.place(x=135, y=85)
day_label_4.place(x=135, y=125)
day_label_5.place(x=135, y=165)
day_label_6.place(x=135, y=205)
day_label_7.place(x=135, y=245)
day_label_8.place(x=135, y=285)

material_number_combobox_1.grid(row=0, column=4, padx=6, pady=4)
material_number_combobox_2.grid(row=1, column=4, padx=6, pady=4)
material_number_combobox_3.grid(row=2, column=4, padx=6, pady=4)
material_number_combobox_4.grid(row=3, column=4, padx=6, pady=4)
material_number_combobox_5.grid(row=4, column=4, padx=6, pady=4)
material_number_combobox_6.grid(row=5, column=4, padx=6, pady=4)
material_number_combobox_7.grid(row=6, column=4, padx=6, pady=4)
material_number_combobox_8.grid(row=7, column=4, padx=6, pady=4)

layout_spacer_label.grid(row=7, column=5, padx=15, pady=4)

row_copybutton_0['state'] = "disabled"
row_copybutton_0.place(x=298,y=19 ,width = 20, height = 20)
row_copybutton_1.place(x=298,y=59 ,width = 20, height = 20)
row_copybutton_2.place(x=298,y=99 ,width = 20, height = 20)
row_copybutton_3.place(x=298,y=139 ,width = 20, height = 20)
row_copybutton_4.place(x=298,y=179 ,width = 20, height = 20)
row_copybutton_5.place(x=298,y=219 ,width = 20, height = 20)
row_copybutton_6.place(x=298,y=259 ,width = 20, height = 20)
row_copybutton_7.place(x=298,y=299 ,width = 20, height = 20)

material_search_checkbutton1.place(x=298, y=0)
material_search_checkbutton2.place(x=298, y=40)
material_search_checkbutton3.place(x=298, y=80)
material_search_checkbutton4.place(x=298, y=120)
material_search_checkbutton5.place(x=298, y=160)
material_search_checkbutton6.place(x=298, y=200)
material_search_checkbutton7.place(x=298, y=240)
material_search_checkbutton8.place(x=298, y=280)

work_class_num_combobox_1.grid(row=0, column=6, padx=0, pady=4)
work_class_num_combobox_2.grid(row=1, column=6, padx=0, pady=4)
work_class_num_combobox_3.grid(row=2, column=6, padx=0, pady=4)
work_class_num_combobox_4.grid(row=3, column=6, padx=0, pady=4)
work_class_num_combobox_5.grid(row=4, column=6, padx=0, pady=4)
work_class_num_combobox_6.grid(row=5, column=6, padx=0, pady=4)
work_class_num_combobox_7.grid(row=6, column=6, padx=0, pady=4)
work_class_num_combobox_8.grid(row=7, column=6, padx=0, pady=4)

work_class_detail_entry_1.grid(row=0, column=7, padx=14, pady=4)
work_class_detail_entry_2.grid(row=1, column=7, padx=14, pady=4)
work_class_detail_entry_3.grid(row=2, column=7, padx=14, pady=4)
work_class_detail_entry_4.grid(row=3, column=7, padx=14, pady=4)
work_class_detail_entry_5.grid(row=4, column=7, padx=14, pady=4)
work_class_detail_entry_6.grid(row=5, column=7, padx=14, pady=4)
work_class_detail_entry_7.grid(row=6, column=7, padx=14, pady=4)
work_class_detail_entry_8.grid(row=7, column=7, padx=14, pady=4)

work_num_combobox_1.grid(padx=10, pady=4)
work_num_combobox_2.grid(padx=10, pady=4)
work_num_combobox_3.grid(padx=10, pady=4)
work_num_combobox_4.grid(padx=10, pady=4)
work_num_combobox_5.grid(padx=10, pady=4)
work_num_combobox_6.grid(padx=10, pady=4)
work_num_combobox_7.grid(padx=10, pady=4)
work_num_combobox_8.grid(padx=10, pady=4)

time_h_num_combobox_1.grid(pady=4)
time_h_num_combobox_2.grid(pady=4)
time_h_num_combobox_3.grid(pady=4)
time_h_num_combobox_4.grid(pady=4)
time_h_num_combobox_5.grid(pady=4)
time_h_num_combobox_6.grid(pady=4)
time_h_num_combobox_7.grid(pady=4)
time_h_num_combobox_8.grid(pady=4)


time_m_num_combobox_1.grid(pady=4)
time_m_num_combobox_2.grid(pady=4)
time_m_num_combobox_3.grid(pady=4)
time_m_num_combobox_4.grid(pady=4)
time_m_num_combobox_5.grid(pady=4)
time_m_num_combobox_6.grid(pady=4)
time_m_num_combobox_7.grid(pady=4)
time_m_num_combobox_8.grid(pady=4)

endtime_h_num_combobox_1.grid(pady=4)
endtime_h_num_combobox_2.grid(pady=4)
endtime_h_num_combobox_3.grid(pady=4)
endtime_h_num_combobox_4.grid(pady=4)
endtime_h_num_combobox_5.grid(pady=4)
endtime_h_num_combobox_6.grid(pady=4)
endtime_h_num_combobox_7.grid(pady=4)
endtime_h_num_combobox_8.grid(pady=4)

endtime_m_num_combobox_1.grid(pady=4)
endtime_m_num_combobox_2.grid(pady=4)
endtime_m_num_combobox_3.grid(pady=4)
endtime_m_num_combobox_4.grid(pady=4)
endtime_m_num_combobox_5.grid(pady=4)
endtime_m_num_combobox_6.grid(pady=4)
endtime_m_num_combobox_7.grid(pady=4)
endtime_m_num_combobox_8.grid(pady=4)

break_listbox_1.grid(row=0, column=1,padx=5, pady=4)
break_listbox_2.grid(row=1, column=1,padx=5, pady=4)
break_listbox_3.grid(row=2, column=1,padx=5, pady=4)
break_listbox_4.grid(row=3, column=1,padx=5, pady=4)
break_listbox_5.grid(row=4, column=1,padx=5, pady=4)
break_listbox_6.grid(row=5, column=1,padx=5, pady=4)
break_listbox_7.grid(row=6, column=1,padx=5, pady=4)
break_listbox_8.grid(row=7, column=1,padx=5, pady=4)

row_clear_button_1.grid(row=0, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_2.grid(row=1, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_3.grid(row=2, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_4.grid(row=3, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_5.grid(row=4, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_6.grid(row=5, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_7.grid(row=6, column=2, padx=1, pady=4, ipadx=17)
row_clear_button_8.grid(row=7, column=2, padx=1, pady=4, ipadx=17)

total_label_1.grid(pady=6, padx=1)
total_label_2.grid(pady=6, padx=1)
total_label_3.grid(pady=6, padx=1)
total_label_4.grid(pady=6, padx=1)
total_label_5.grid(pady=6, padx=1)
total_label_6.grid(pady=6, padx=1)
total_label_7.grid(pady=6, padx=1)
total_label_8.grid(pady=6, padx=1)


total_time_min.place(x=10,y=5)
overtime_time_min.place(x=230,y=5)
nightshift_minutes_label.place(x=450,y=5)
set_button.pack(padx=100, ipadx=100, ipady=10)





root1.resizable(width=False, height=False)
root1.mainloop()




