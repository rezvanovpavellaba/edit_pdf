

    

    



def Page2():

    import streamlit as st
    import fitz  # PyMuPDF
    import pandas as pd
    from io import BytesIO

    def redact_text_on_page(page, df, page_number):
        """
        Заменяет текст на указанной странице документа.

        :param page: Объект страницы документа
        :param df: DataFrame с данными для замены
        :param page_number: Номер страницы для обработки
        """
        df['Page'] = df['Page'].astype(str)

        # Фильтруем данные для текущей страницы
        df_page = df[df['Page'] == str(page_number)]

        for i, raw_text in enumerate(df_page['Old Value'].values):
            #сам поиск текста по старым значениям из Excel, который нужно заменить

            hits = page.search_for(raw_text)

            new_text = df_page['New Value'].values[i]
            
            # Параметры для редактирования
            new_fontsize = 7.86  # Новый размер шрифта
            new_width = -0.1    # Новая ширина прямоугольника
            new_width_2 = -14   # Новая ширина прямоугольника
            new_height = -1.1   # Новая высота прямоугольника
            new_height_2 = -0.85   # Новая высота прямоугольника

            for rect in hits:
                x1, y1, x2, y2 = rect
                new_x1 = x1 + new_width_2
                new_x2 = x2 + new_width
                new_y2 = y2 - new_height_2*1.1
                new_y1 = y1 + new_height
                new_rect = fitz.Rect(new_x1,new_y1, new_x2, new_y2)

                # Добавляем аннотацию для редактирования
                page.add_redact_annot(new_rect, new_text,
                                    fontsize=new_fontsize,
                                    fontname=page.get_fonts()[1][4],
                                    align=fitz.TEXT_ALIGN_RIGHT)

            # Применяем редактирование
            page.apply_redactions()

    def process_pdf(pdf_file, excel_data, sheet_name):
        """Обрабатывает PDF файл, редактируя текст на основе данных Excel."""
        pdf_bytes = BytesIO(pdf_file.read())
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")

        sheet_data = excel_data[sheet_name]
        for page_number in range(len(doc)):
            redact_text_on_page(doc[page_number], sheet_data, page_number)

        output = BytesIO()
        doc.save(output)
        doc.close()
        return output

    # Streamlit UI
    st.title("PDF и Excel обработчик для редактирования")

    uploaded_pdfs = st.file_uploader("Загрузите PDF файлы", type="pdf", accept_multiple_files=True)
    uploaded_excel = st.file_uploader("Загрузите Excel файл", type="xlsx")

    if uploaded_pdfs and uploaded_excel:
        # Чтение Excel файла
        excel_data = pd.read_excel(uploaded_excel, sheet_name=None, dtype=str)

        processed_files = []
        for pdf_file in uploaded_pdfs:
            # Убираем расширение у имени PDF файла
            pdf_name = pdf_file.name.rsplit('.', 1)[0]
            

            # Проверяем совпадение между именем PDF и листами Excel
            if pdf_name in excel_data:
                processed_pdf = process_pdf(pdf_file, excel_data, pdf_name)
                processed_files.append((pdf_file.name, processed_pdf))
            else:
                st.warning(f"Лист Excel для файла {pdf_file.name} не найден.")

        if processed_files:
            for name, processed_pdf in processed_files:
                st.download_button(
                    label=f"Скачать обработанный {name}",
                    data=processed_pdf.getvalue(),
                    file_name=f"updated_{name}",
                    mime="application/pdf",
                )
        else:
            st.warning("Не найдено совпадений между файлами PDF и листами Excel.")
