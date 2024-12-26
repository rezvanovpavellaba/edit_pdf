def Page1():

    import streamlit as st
    import pandas as pd
    import xml.etree.ElementTree as ET
    from io import BytesIO
    import json
    import os
    from st_aggrid import AgGrid, GridOptionsBuilder, DataReturnMode
    from datetime import datetime, timedelta

    # Инициализация состояния для excel_sheets
    if "excel_sheets" not in st.session_state:
        st.session_state["excel_sheets"] = {}

    # Функция для парсинга XML
    def parse_xml(file):
        tree = ET.parse(file)
        root = tree.getroot()

        data = []
        for sample in root.findall(".//SAMPLE"):
            sample_name = sample.attrib.get("name", "")
            create_time = sample.attrib.get("createtime", "")
            for compound in sample.findall("COMPOUND"):
                compound_name = compound.attrib.get("name", "")
                for peak in compound.findall("PEAK"):
                    response = peak.attrib.get("response", "")
                    conc = peak.attrib.get("analconc", "")
                    data.append({
                        "Sample_name": sample_name,
                        "Create_time": create_time,
                        "Compound_name": compound_name,
                        "Response": response,
                        "Conc": conc
                    })

        df = pd.DataFrame(data)

        df.sort_values(by=["Compound_name", "Sample_name"], inplace=True)
        return df


    # Функция для загрузки JSON
    def load_json(file):
        return pd.DataFrame(json.load(file))

    # Функция для извлечения кривых из XML
    def extract_curves(file):
        try:
            file.seek(0)
            tree = ET.parse(file)
            root = tree.getroot()

            data = {}
            for compound in root.findall(".//CALIBRATIONDATA/COMPOUND"):
                compound_name = compound.attrib.get("name", "")
                calibration_curve = compound.find(".//CALIBRATIONCURVE")
                if calibration_curve is not None:
                    curve = calibration_curve.attrib.get("curve", "")
                    data[compound_name] = curve

            return data
        except ET.ParseError as e:
            raise ValueError(f"XML parsing error: {e}")
        except Exception as e:
            raise ValueError(f"Unexpected error: {e}")

    # Функция для вычисления результата
    def calculate_result(new_conc, curve):
        try:
            x = float(new_conc)
            result = eval(curve.replace("x", str(x)))
            return result
        except Exception:
            return ""

    # Функция для округления значений
    def apply_rounding(data, precision):
        """Округляет значения в массиве данных и обеспечивает единообразный формат вывода."""
        def round_number(x):
            try:
                num = float(x)
                return round(num, precision)
            except ValueError:
                return x

        return [round_number(x) for x in data]

    # Функция для применения разрядности (форматирование вывода как в Excel)
    def apply_digits(data, digits):
        """Применяет разрядность чисел, добавляя фиксированное количество знаков после запятой."""
        def format_number(x):
            try:
                num = float(x)
                return f"{num:.{digits}f}"
            except ValueError:
                return x

        return [format_number(x) for x in data]
    
    # Streamlit приложение
    st.title("XML & JSON Integration for Excel Export")

    # Выбор режима работы
    mode = st.radio("Select mode:", ("Use curves from XML", "Manually specify coefficients"))

    uploaded_xml_files = st.file_uploader("Upload XML Files", type="xml", accept_multiple_files=True)
    file_data = {}
    curves_dict = {}
    coefficients = {}
    compounds = set()
    excel_sheets = {}

    if uploaded_xml_files:
        for file in uploaded_xml_files:
            try:
                df = parse_xml(file)
                file_data[file.name] = df
            except Exception as e:
                st.error(f"Error processing XML file {file.name}: {e}")

            try:
                curves = extract_curves(file)
                curves_dict.update(curves)
            except ValueError as e:
                st.error(f"Error extracting curves from XML file {file.name}: {e}")
            except Exception as e:
                st.error(f"Unexpected error with file {file.name}: {e}")

        for df in file_data.values():
            compounds.update(df["Compound_name"].unique())

        if mode == "Manually specify coefficients":
            with st.expander("Specify coefficients for each compound"):
               for compound in sorted(compounds):
                   # Уникальные ключи для сессии
                   key_a = f"a_{compound}"
                   key_b = f"b_{compound}"
                   
                   # Инициализация значений, если они еще не сохранены
                   if key_a not in st.session_state:
                       st.session_state[key_a] = 0.0
                   if key_b not in st.session_state:
                       st.session_state[key_b] = 0.0

                   # Создание виджетов и сохранение значений в сессии
                   a = st.number_input(
                       f"Coefficient a for {compound}", 
                       value=st.session_state[key_a], 
                       key=f"key_{key_a}"
                   )
                   b = st.number_input(
                       f"Coefficient b for {compound}", 
                       value=st.session_state[key_b], 
                       key=f"key_{key_b}"
                   )

                   # Обновление значений в сессии
                   st.session_state[key_a] = a
                   st.session_state[key_b] = b

                   # Сохранение в словарь coefficients
                   coefficients[compound] = (a, b)

        st.subheader("Upload JSON Files for Each Compound")
        json_files = {}
        for compound in sorted(compounds):
            uploaded_json = st.file_uploader(f"Upload JSON for Compound: {compound}", type="json", key=compound)
            if uploaded_json:
                try:
                    json_files[compound] = load_json(uploaded_json)
                except Exception as e:
                    st.error(f"Error processing JSON for {compound}: {e}")

        for compound in sorted(compounds):
            # Инициализация сессии для округления
            for col in ["response", "conc", "new_response", "new_conc"]:
                if f"{col}_rounding_{compound}" not in st.session_state:
                    st.session_state[f"{col}_rounding_{compound}"] = 2  # Значение по умолчанию
                if f"{col}_digits_{compound}" not in st.session_state:
                    st.session_state[f"{col}_digits_{compound}"] = 2  # Значение по умолчанию
            
            with st.sidebar:
                 # Динамическое создание виджетов
                 with st.expander(f"Настройка округления и разрядности для {compound}"):
                      # Виджеты для округления
                      response_rounding = st.number_input(
                          f"Округление для Response ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"response_rounding_{compound}"],
                          step=1, key=f"key_response_rounding_{compound}"
                      )
                      conc_rounding = st.number_input(
                          f"Округление для Conc ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"conc_rounding_{compound}"],
                          step=1, key=f"key_conc_rounding_{compound}"
                      )
                      new_response_rounding = st.number_input(
                          f"Округление для newResponse ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"new_response_rounding_{compound}"],
                          step=1, key=f"key_new_response_rounding_{compound}"
                      )
                      new_conc_rounding = st.number_input(
                          f"Округление для newConc ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"new_conc_rounding_{compound}"],
                          step=1, key=f"key_new_conc_rounding_{compound}"
                      )

                      # Виджеты для разрядности
                      response_digits = st.number_input(
                          f"Разрядность для Response ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"response_digits_{compound}"],
                          step=1, key=f"key_response_digits_{compound}"
                      )
                      conc_digits = st.number_input(
                          f"Разрядность для Conc ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"conc_digits_{compound}"],
                          step=1, key=f"key_conc_digits_{compound}"
                      )
                      new_response_digits = st.number_input(
                          f"Разрядность для newResponse ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"new_response_digits_{compound}"],
                          step=1, key=f"key_new_response_digits_{compound}"
                      )
                      new_conc_digits = st.number_input(
                          f"Разрядность для newConc ({compound})", min_value=0, max_value=10,
                          value=st.session_state[f"new_conc_digits_{compound}"],
                          step=1, key=f"key_new_conc_digits_{compound}"
                      )

                      # Обновляем значения в сессии
                      st.session_state[f"response_rounding_{compound}"] = response_rounding
                      st.session_state[f"conc_rounding_{compound}"] = conc_rounding
                      st.session_state[f"new_response_rounding_{compound}"] = new_response_rounding
                      st.session_state[f"new_conc_rounding_{compound}"] = new_conc_rounding

                      st.session_state[f"response_digits_{compound}"] = response_digits
                      st.session_state[f"conc_digits_{compound}"] = conc_digits
                      st.session_state[f"new_response_digits_{compound}"] = new_response_digits
                      st.session_state[f"new_conc_digits_{compound}"] = new_conc_digits


        if st.button("Process and Export to Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for file_name, df in file_data.items():
                    for compound, json_df in json_files.items():
                        compound_df = df[df["Compound_name"] == compound].copy()

                        if mode == "Use curves from XML":
                            curve = curves_dict.get(compound, "")
                        else:
                            a, b = coefficients.get(compound, (0, 0))
                            curve = f"{a}*x+{b}"

                        json_df["Identifier"] = json_df.apply(
                            lambda x: f"-{x['Subject']}-{x['Period']}-{str(int(x['timePoint']) - 1).zfill(2)}",
                            axis=1
                        )
                        id_to_conc = dict(zip(json_df["Identifier"], json_df["CalcConc"]))

                        compound_df["newConc"] = compound_df["Sample_name"].apply(
                            lambda x: next((value for key, value in id_to_conc.items() if x.endswith(key)), "")
                        )

                        compound_df["newConc"] = compound_df["newConc"].apply(lambda x: "" if x == 0 else x)
                        compound_df["newResponse"] = compound_df.apply(
                            lambda row: calculate_result(row["newConc"], curve) if row["newConc"] else "",
                            axis=1
                        )

                        df.loc[df["Compound_name"] == compound, "newConc"] = compound_df["newConc"]
                        df.loc[df["Compound_name"] == compound, "newResponse"] = compound_df["newResponse"]
                    
                    # Применяем округление отдельно для каждой группы Compound_name
                    for compound in df["Compound_name"].unique():

                        # Получаем настройки округления и разрядности из сессии
                        conc_rounding = st.session_state.get(f"conc_rounding_{compound}", 2)
                        response_rounding = st.session_state.get(f"response_rounding_{compound}", 2)
                        conc_digits = st.session_state.get(f"conc_digits_{compound}", 2)
                        response_digits = st.session_state.get(f"response_digits_{compound}", 2)

                        # Применяем округление к столбцам "Conc" и "Response"
                        if "Conc" in df.columns:
                            df.loc[df["Compound_name"] == compound, "Conc"] = apply_rounding(
                                df.loc[df["Compound_name"] == compound, "Conc"], conc_rounding
                            )
                            df.loc[df["Compound_name"] == compound, "Conc"] = apply_digits(
                                df.loc[df["Compound_name"] == compound, "Conc"], conc_digits
                            )
                        else:
                            st.warning(f"Column 'Conc' not found for compound {compound}.")

                        if "Response" in df.columns:
                            df.loc[df["Compound_name"] == compound, "Response"] = apply_rounding(
                                df.loc[df["Compound_name"] == compound, "Response"], response_rounding
                            )
                            df.loc[df["Compound_name"] == compound, "Response"] = apply_digits(
                                df.loc[df["Compound_name"] == compound, "Response"], response_digits
                            )
                        else:
                            st.warning(f"Column 'Response' not found for compound {compound}.")

                        new_conc_rounding = st.session_state.get(f"new_conc_rounding_{compound}", 2)  # Значение по умолчанию
                        new_response_rounding = st.session_state.get(f"new_response_rounding_{compound}", 2)
                        new_conc_digits = st.session_state.get(f"new_conc_digits_{compound}", 2)
                        new_response_digits = st.session_state.get(f"new_response_digits_{compound}", 2)

                        if "newConc" in df.columns:
                            df.loc[df["Compound_name"] == compound, "newConc"] = apply_rounding(
                                df.loc[df["Compound_name"] == compound, "newConc"],
                                new_conc_rounding
                            )
                            df.loc[df["Compound_name"] == compound, "newConc"] = apply_digits(
                                                    df.loc[df["Compound_name"] == compound, "newConc"], new_conc_digits
                                                )
                        else:
                            st.warning(f"Column 'newConc' not found for compound {compound}.")

                        if "newResponse" in df.columns:
                            df.loc[df["Compound_name"] == compound, "newResponse"] = apply_rounding(
                                df.loc[df["Compound_name"] == compound, "newResponse"],
                                new_response_rounding
                            )
                            df.loc[df["Compound_name"] == compound, "newResponse"] = apply_digits(
                                df.loc[df["Compound_name"] == compound, "newResponse"], new_response_digits
                            )
                        else:
                            st.warning(f"Column 'newResponse' not found for compound {compound}.")

                        

                    sheet_name = file_name.replace("/", "_").replace("\\", "_")[:31]

                    # Добавляем новую пустую колонку newCreate_time
                    df["newCreate_time"] = [""] * len(df)

                    # Изменение порядка колонок
                    desired_column_order = [
                        "Compound_name", "Sample_name", "Create_time", "newCreate_time",
                        "Response", "newResponse", "Conc", "newConc"
                    ]
                    existing_columns = [col for col in desired_column_order if col in df.columns]
                    df = df[existing_columns]

                    # Сохранение данных в Excel
                    df.to_excel(writer, index=False, sheet_name=sheet_name)


                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                    excel_sheets[sheet_name] = df  # Сохраняем данные для визуализации

                    st.session_state["excel_sheets"] = excel_sheets

                    worksheet = writer.sheets[sheet_name]
                    for i, column in enumerate(df.columns):
                        column_width = max(df[column].astype(str).map(len).max(), len(column)) + 2
                        worksheet.set_column(i, i, column_width)

            output.seek(0)
            st.download_button(
                label="Download Excel File",
                data=output,
                file_name="processed_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # Визуализация и редактирование таблиц по листам
        if "excel_sheets" in st.session_state and st.session_state["excel_sheets"]:
            st.subheader("View and Edit Excel Sheets")

            # Выбор текущего листа
            selected_sheet = st.selectbox("Select Sheet to View/Edit:", options=list(st.session_state["excel_sheets"].keys()))


            # Получаем данные для текущего листа
            df = st.session_state["excel_sheets"][selected_sheet]

            # Настройка параметров AgGrid
            gb = GridOptionsBuilder.from_dataframe(df)
            gb.configure_default_column(editable=True)  # Делаем все колонки редактируемыми
            gb.configure_grid_options(enableRangeSelection=True)
            gb.configure_grid_options(enableFullScreen=True)  # Включаем полноэкранный режим
            grid_options = gb.build()

            # Отображаем таблицу с возможностью редактирования
            grid_response = AgGrid(
                df,
                gridOptions=grid_options,
                data_return_mode=DataReturnMode.FILTERED_AND_SORTED,
                update_mode="MODEL_CHANGED",  # Автоматическое обновление данных при изменении
                fit_columns_on_grid_load=True,
                enable_enterprise_modules=False,
                editable=True,
            )

            col1, col2 = st.columns(2)

            with col1:
                 #Обновление страницы путем ререндинга
                 if st.button("Update the table"):
                    st.query_params = {"rerun": "true"}

            updated_df = pd.DataFrame(grid_response["data"])
            
            with col2:
                #Очистка данных колонки newCreate_time
                if st.button("Clear 'newCreate_time' column"):
                   if "newCreate_time" in updated_df.columns:
                       updated_df["newCreate_time"] = ""  # Очищаем колонку
                       st.session_state["excel_sheets"][selected_sheet] = updated_df
                       st.success("'newCreate_time' column has been cleared successfully!")
                   else:
                       st.warning("'newCreate_time' column does not exist in the selected sheet.")
            
            # Обновляем данные в сессии
            st.session_state["excel_sheets"][selected_sheet] = updated_df
            st.success(f"Changes to sheet '{selected_sheet}' saved successfully!")

            updated_df = st.session_state["excel_sheets"][selected_sheet].reset_index(drop=True)

            # Проверяем наличие колонок Create_time и newCreate_time
            if "Create_time" in updated_df.columns and "newCreate_time" in updated_df.columns:
                try:
                    # Проверяем, что значение в первой строке newCreate_time корректное

                    first_value = updated_df.loc[0, "newCreate_time"]

                    # Проверяем, не пустое ли значение и строковый ли тип
                    if isinstance(first_value, str) and len(first_value.split(":")) == 3:
                        first_time = datetime.strptime(first_value, "%H:%M:%S")
                        first_hour = first_time.hour
                    else:
                        st.warning("The first value of 'newCreate_time' is not in a valid time format (HH:MM:SS). Update skipped.")
                        first_time = None

                    # Выполняем обновление, только если значение корректное
                    if first_time:
                        for i in range(1, len(updated_df)):
                            try:
                                # Преобразуем значения Create_time в datetime
                                prev_time = datetime.strptime(str(updated_df.loc[i - 1, "Create_time"]), "%H:%M:%S")
                                curr_time = datetime.strptime(str(updated_df.loc[i, "Create_time"]), "%H:%M:%S")

                                # Рассчитываем разницу времени
                                time_difference = curr_time - prev_time 

                                # Изменяем час в prev_time на hour из first_time
                                prev_time = prev_time.replace(hour=first_hour)

                                new_time = prev_time + time_difference
                                new_time = new_time.time()
                                
                                # Записываем новое время в newCreate_time
                                updated_df.loc[i, "newCreate_time"] = new_time.strftime("%H:%M:%S")

                                
                            except ValueError as row_error:
                                st.warning(f"Row {i} has invalid time format in 'Create_time': {row_error}")
                            except Exception as row_error:
                                st.warning(f"Unexpected error processing row {i}: {row_error}")

                        st.session_state["excel_sheets"][selected_sheet] = updated_df

                except Exception as e:
                    st.error(f"Error while processing newCreate_time: {e}")
            else:
                st.warning("Columns 'Create_time' and/or 'newCreate_time' are missing. No updates applied.")

        

        # Инициализация сессии для хранения значений виджетов ввода
        if "page_inputs" not in st.session_state:
            st.session_state["page_inputs"] = {}

        # Динамическое создание виджетов ввода для Page
        if st.session_state["excel_sheets"]:
            st.subheader("Generate New Excel File with Old and New Values")

            # Создание словаря для хранения значений Page для каждого соединения
            page_inputs = st.session_state["page_inputs"]

            with st.sidebar:
                 # Динамическое создание виджетов
                 with st.expander("Номера страниц"):
                      for sheet_name, df in st.session_state["excel_sheets"].items():
                          st.write(f"### Sheet: {sheet_name}")
                          compounds = df["Compound_name"].unique()

                          
                          for compound in compounds:
                              st.write(f"#### Compound: {compound}")
                              
                              
                              # Уникальные ключи для виджетов
                              key_newConc_newResponse  = f"newConc_newResponse _{sheet_name}_{compound}_value"
                              key_newCreate_time = f"newCreate_time_{sheet_name}_{compound}_value"
                              
                              # Инициализация значений виджетов
                              if key_newConc_newResponse not in st.session_state:
                                  st.session_state[key_newConc_newResponse] = 1
                              if key_newCreate_time not in st.session_state:
                                  st.session_state[key_newCreate_time] = 0

                              st.session_state[key_newConc_newResponse] = st.number_input(
                                  f"Page for newConc and newResponse ({compound})", min_value=0, step=1, value=st.session_state[key_newConc_newResponse], key=f"key_{key_newConc_newResponse}"
                              )
                          
                              st.session_state[key_newCreate_time] = st.number_input(
                                  f"Page for newCreate_time ({compound})", min_value=0, step=1, value=st.session_state[key_newCreate_time], key=f"key_{key_newCreate_time}"
                              )

                              # Сохраняем значения в сессию
                              page_inputs[(sheet_name, compound)] = {
                                  "newConc_newResponse_page": st.session_state[key_newConc_newResponse],
                                  "newCreate_time_page": st.session_state[key_newCreate_time],
                              }

            # Генерация нового файла Excel
            if st.button("Generate New Excel with Old and New Values"):
                new_output = BytesIO()
                with pd.ExcelWriter(new_output, engine='xlsxwriter') as new_writer:
                    
                    workbook = new_writer.book  # Получаем доступ к книге Excel
                    text_format = workbook.add_format({'num_format': '@'})  # Формат для текста

                    for sheet_name, df in st.session_state["excel_sheets"].items():
                        
                            # Фильтруем строки, где newConc и newResponse не пусты
                            filtered_df = df[
                                (df["newConc"].notna()) & (df["newResponse"].notna())
                            ]

                            if not filtered_df.empty:
                                # Создаём DataFrame для Old Value, New Value и Page
                                result_df = pd.DataFrame({
                                    "Old Value": list(filtered_df["Response"]) + list(filtered_df["Conc"]) + list(filtered_df["Create_time"]),
                                    "New Value": list(filtered_df["newResponse"]) + list(filtered_df["newConc"]) + list(filtered_df["newCreate_time"]),
                                    "Page": sum([
                                        [
                                            page_inputs[(sheet_name, row["Compound_name"])][f"{key}_page"]
                                            for _, row in filtered_df.iterrows()
                                        ]
                                        for key in ["newConc_newResponse","newConc_newResponse", "newCreate_time"]
                                    ], [])
                                })

                                # Удаляем строки, где New Value пустое
                                result_df = result_df[
                                    result_df["New Value"].notna()
                                    & (result_df["New Value"] != "")
                                    & (result_df["New Value"].astype(str).str.strip() != "")
                                ]

                                # Удаляем строки, где Old Value пустое
                                result_df = result_df[
                                    result_df["Old Value"].notna()
                                    & (result_df["Old Value"] != "")
                                    & (result_df["Old Value"].astype(str).str.strip() != "")
                                ]

                                # Формируем название листа
                                clean_sheet_name = f"{sheet_name.replace('.xml', '').replace('.', '_')}"[:31]  # Ограничение по длине имени листа

                                # Записываем результат на новый лист
                                result_df.to_excel(new_writer, index=False, sheet_name=clean_sheet_name)

                                worksheet = new_writer.sheets[clean_sheet_name]  # Получаем доступ к текущему листу
                    
                                # Применяем текстовый формат ко всем ячейкам
                                for col_num, _ in enumerate(result_df.columns):
                                    worksheet.set_column(col_num, col_num, None, text_format)

                new_output.seek(0)
                st.download_button(
                    label="Download Filtered Excel File",
                    data=new_output,
                    file_name="filtered_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.success("New Excel file generated successfully!")

    