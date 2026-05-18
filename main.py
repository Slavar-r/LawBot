import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api.upload import VkUpload
import sqlite3
import json
import os
from datetime import datetime
from docxtpl import DocxTemplate

"""
VK_TOKEN - Вставляете сгенерированный ключ от вк (Надо создать сообщество). 
    В его настройках разрешить управление сообществом, сообщениями и документами (Раздел работа с API)
В переменную MAIN_PATH Указывается путь до каталога с папками шаблонов
Устанавливаете зависимости: в командной строке(cmd) пишете (pip install <Название пакета>). Пакеты: vk-api doxtpl
Запускаете скрипт 
Пишете боту "начать" или "новый" и активируется скрипт
"""
# -------------- НАСТРОЙКИ -----------------
VK_TOKEN = "vk1.a.QgR6sx9uVSwVjmr0RbxL0gxGE0i4CO8p1as7wO6_ozAIe2W_hhA354x9hcKQUW1s6qJYwcEMU_cNoI4IG4GxeyLNTF3q4RFojK_QYx1HJquE-yv7zKHJHyDatKjuEZyzAlmop4I7fwryS1Hg4ZtlXO-5G4mEr8eLmE7n6U4htNbRAM09CLo0rdN7plpxDLbzEjoSUtPI63vMtSECUkdxuA"
MAIN_PATH= r"C:/Users/user/Desktop/AlsuProg/Doc_sample"
SAMPLES=dict({
              "Об увольнении":"Firing","О трудоустройстве":"GetWork","Об отпуске":"Vacancy","Извещение об отпуске":"VacancyNotify","Регламент отпуска":"VacancyNotifyQuest",
              "Увольнение":"Firing","Трудоустройство":"GetWork","Отпуск":"Vacancy","Оповещение об отпуске":"VacancyNotify","Вопрос оповещения об отпуске":"VacancyNotifyQuest"
            })
DocType_Rus=dict({
                  "Firing":"Заявление на увольнение","Об увольнении":"Заявление на увольнение",
                  "GetWork":"Заявление на трудоустройство","О трудоустройстве":"Заявление на трудоустройство",
                  "Vacancy":"Заявление на отпуск","Об отпуске":"Заявление на отпуск",
                  "VN":"Оповещение об отпуске","VacancyNotify":"Оповещение об отпуске","Извещение об отпуске":"Оповещение об отпуске",
                  "VNC":"Уведомление об отпуске", "VacancyNotifyQuest":"Уведомление об отпуске","Регламент отпуска":"Уведомление об отпуске"
                })
DOCS_VARIABLES = dict({
                       "Тип":"Тип документа",
                       "Получатель":"ФИО получателя(Кому)","Отправитель":"ФИО отправителя(От кого)",
                       "ДолжностьПолучателя":"Должность(Кому)","ДолжностьОтправителя":"Должность(От кого)",
                       "Организация":"Название организации","Дата":"Дата",
                       "Причина":"Причина переноса", "Отдел":"Отдел сотрудника",
                       "Цена":"Размер оклада", "Кадровик":"ФИО сотрудника отдела кадров",
                       "ДатаРасторжения":"Дата расторжения","ДатаПриемаНаРаботу":"Дата устройства на работу",
                       "ИспытательныйСрок":"Продолжительность испытательного срока","ЖелаемаяДолжность":"Желаемая должность",
                       "ПланНачОтдых":"Начало отпуска до переноса","ПланКонОтдых":"Конец отпуска до переноса",
                       "НачОтдых":"Начало отпуска", "КонОтдых":"Конец отпуска",
                       "КолВоДней":"Кол-во дней в отпуске","ОзнакУвед":"Дата ознакомления с уведомлением",
                     })

DOCUMENT_VARIABLES = dict({
                           "Получатель":"RECIPIENT","Отправитель":"SENDER",
                           "ДолжностьПолучателя":"RECEPIENT_JOB","ДолжностьОтправителя":"SENDER_JOB",
                           "Организация":"ORGANIZATION","Дата":"TIME",
                           "Причина":"CAUSE", "Отдел":"DEPARTMENT",
                           "Цена":"PRICE", "Кадровик":"EMPLOYEE",
                           "ДатаРасторжения":"DESTINATIONTIME","ДатаПриемаНаРаботу":"WORKTIME",
                           "ИспытательныйСрок":"TRIAL","ЖелаемаяДолжность":"OFFER",
                           "ПланНачОтдых":"STARTPLANVACANCY","ПланКонОтдых":"ENDPLANVACANCY",
                           "НачОтдых":"STARTVACANCY", "КонОтдых":"ENDVACANCY",
                           "КолВоДней":"NUMVACDAYS","ОзнакУвед":"TIMECONFIRM",
                           
                           
                         })
SCHEMA_DOC = {
    "Об увольнении":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["ДолжностьОтправителя"],DOCS_VARIABLES["Отправитель"],DOCS_VARIABLES["ДатаРасторжения"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["Кадровик"]],
    "О трудоустройстве":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["ДолжностьОтправителя"],DOCS_VARIABLES["Отправитель"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["Цена"],DOCS_VARIABLES["ЖелаемаяДолжность"],DOCS_VARIABLES["ИспытательныйСрок"],DOCS_VARIABLES["ДатаПриемаНаРаботу"]],
    "Об отпуске":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["ДолжностьОтправителя"],DOCS_VARIABLES["Отправитель"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["Причина"],DOCS_VARIABLES["ПланНачОтдых"],DOCS_VARIABLES["ПланКонОтдых"],DOCS_VARIABLES["НачОтдых"],DOCS_VARIABLES["КонОтдых"],DOCS_VARIABLES["КолВоДней"]],
    "Извещение об отпуске":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["НачОтдых"],DOCS_VARIABLES["КонОтдых"],DOCS_VARIABLES["Кадровик"], DOCS_VARIABLES["КолВоДней"]],
    "Регламент отпуска":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["ОзнакУвед"],DOCS_VARIABLES["Кадровик"]]
              }
                       

# ================================
def init_db():
    conn = sqlite3.connect('users_data.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            doc JSON,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def save_document(user_id, data_dict):
    conn = sqlite3.connect('users_data.db')
    c = conn.cursor()
    c.execute("INSERT INTO documents (user_id, doc) VALUES (?, ?)",
              (user_id, json.dumps(data_dict, ensure_ascii=False)))
    conn.commit()
    doc_id = c.lastrowid
    conn.close()
    return doc_id

def get_document(doc_id):
    conn = sqlite3.connect('users_data.db')
    c = conn.cursor()
    c.execute("SELECT doc, created_at FROM documents WHERE id = ?", (doc_id,))
    row = c.fetchone()
    conn.close()
    if row:
        doc = json.loads(row[0])
        doc['created_at'] = row[1]  # добавляем дату создания в словарь
        #print(doc)
        return doc
    return None

def update_document(doc_id, new_data):
    conn = sqlite3.connect('users_data.db')
    c = conn.cursor()
    c.execute("UPDATE documents SET doc = ? WHERE id = ?",
              (json.dumps(new_data, ensure_ascii=False), doc_id))
    conn.commit()
    updated = c.rowcount > 0
    conn.close()
    return updated

def get_user_docs(user_id):
    conn = sqlite3.connect('users_data.db')
    c = conn.cursor()
    c.execute("SELECT id, doc FROM documents WHERE user_id = ?", (user_id,))
    rows = c.fetchall()
    conn.close()
    return [(row[0], json.loads(row[1])) for row in rows]

def generate_docx(doc_id, data, DocName):
    SAMPLE_NAME=MAIN_PATH+rf"/{SAMPLES[DocName]}"+ "/"+ SAMPLES[DocName]+"_0.docx"
    if not os.path.exists(SAMPLE_NAME):
        raise FileNotFoundError(f"Шаблон не найден: {SAMPLE_NAME}")
    doc = DocxTemplate(SAMPLE_NAME)
    
    match data["тип_документа"]:
        case "Об увольнении":
            context = {
            "RECEPIENT_JOB":data["Должность(Кому)"],
                "ORGANIZATION":data["Название организации"],
                "RECIPIENT":data["ФИО получателя(Кому)"],
                "SENDER_JOB":data["Должность(От кого)"],
                "DEPARTMENT":data["Отдел сотрудника"],
                "SENDER":data["ФИО отправителя(От кого)"],
                "DESTINATIONTIME":data["Дата расторжения"],
                "EMPLOYEE":data["ФИО сотрудника отдела кадров"],
                "TIME":data["Дата"]
                      }
        case "О трудоустройстве":
            context = {
                "RECEPIENT_JOB":data["Должность(Кому)"],
                "ORGANIZATION":data["Название организации"],
                "RECIPIENT":data["ФИО получателя(Кому)"],
                "SENDER_JOB":data["Должность(От кого)"],
                "DEPARTMENT":data["Отдел сотрудника"],
                "SENDER":data["ФИО отправителя(От кого)"],
                "TIME":data["Дата"],
                "OFFER":data["Желаемая должность"],
                "WORKTIME":data["Дата устройства на работу"],
                "TRIAL":data["Продолжительность испытательного срока"],
                "PRICE":data["Размер оклада"]
                      }
        case "Об отпуске":
            context = {
                "RECEPIENT_JOB":data["Должность(Кому)"],
                "ORGANIZATION":data["Название организации"],
                "SENDER_JOB":data["Должность(От кого)"],
                "DEPARTMENT":data["Отдел сотрудника"],
                "SENDER":data["ФИО отправителя(От кого)"],
                "TIME":data["Дата"],
                "CAUSE":data["Причина переноса"],
                "STARTPLANVACANCY":data["Начало отпуска до переноса"],
                "ENDPLANVACANCY":data["Конец отпуска до переноса"],
                "STARTVACANCY":data["Начало отпуска"],
                "ENDVACANCY":data["Конец отпуска"],
                "NUMVACDAYS":data["Кол-во дней в отпуске"],
                      }
        case "Извещение об отпуске":
            context = {
                "RECEPIENT_JOB":data["Должность(Кому)"],
                "RECIPIENT":data["ФИО получателя(Кому)"],
                "ORGANIZATION":data["Название организации"],
                "DEPARTMENT":data["Отдел сотрудника"],
                "TIME":data["Дата"],
                "STARTVACANCY":data["Начало отпуска"],
                "ENDVACANCY":data["Конец отпуска"],
                "NUMVACDAYS":data["Кол-во дней в отпуске"],
                "EMPLOYEE":data["ФИО сотрудника отдела кадров"],
                     }
        case "Регламент отпуска":
            context = {
                "RECEPIENT_JOB":data["Должность(Кому)"],
                "ORGANIZATION":data["Название организации"],
                "RECIPIENT":data["ФИО получателя(Кому)"],
                "DEPARTMENT":data["Отдел сотрудника"],
                "EMPLOYEE":data["ФИО сотрудника отдела кадров"],
                "TIME":data["Дата"],
                "TIMECONFIRM":data["Дата ознакомления с уведомлением"],
                      }
            
    doc.render(context)
    # Сохраняем файл
    filename = f"{DocType_Rus[DocName]} {doc_id} {datetime.now().strftime('%d.%m.%Y')}.docx"
    filepath = os.path.join(MAIN_PATH+rf"/{SAMPLES[DocName]}", filename)
    doc.save(filepath)
    return filepath

def send_document(user_id, file_path, title="Документ"):
    upload = VkUpload(vk_session)
    try:
        # Загружаем документ на сервер ВК
        doc_upload = upload.document(file_path, title=title, message_peer_id=user_id)
        # Отправляем сообщение с прикреплённым документом
        

        vk.messages.send(
            user_id=user_id,
            message="Документ готов для скачивания",
            attachment=f"doc{doc_upload['doc']['owner_id']}_{doc_upload['doc']['id']}",
            random_id=0
        )
        return True
    except Exception as e:
        print(f"Ошибка отправки документа: {e}")
        return False

#пулинг бота
vk_session = vk_api.VkApi(token=VK_TOKEN)
vk = vk_session.get_api()
longpoll = VkLongPoll(vk_session)

def send_message(user_id, text):
    vk.messages.send(user_id=user_id, message=text, random_id=0)

init_db()
print("Бот запущен")

user_states = {}
for event in longpoll.listen():
    if event.type == VkEventType.MESSAGE_NEW and event.to_me:
        user_id = event.user_id
        text = event.text.strip().lower()

        # Команда: начать создание документа
        if text in ("начать", "новый"):
            user_states[user_id] = {'action': 'choose_type'}
            types_list = "\n".join([f"- {t}" for t in SCHEMA_DOC.keys()])
            send_message(user_id, f"📌 Выберите тему документа:\n{types_list}")
            continue

        # --- Обработка выбора типа ---
        if user_id in user_states and user_states[user_id].get('action') == 'choose_type':
            chosen_type = event.text.strip()#.capitalize()
            if chosen_type in SCHEMA_DOC:
                fields_order = SCHEMA_DOC[chosen_type]
                user_states[user_id] = {
                    'action': 'new_doc',
                    'fields': {'тип_документа': chosen_type},
                    'step': 0,
                    'fields_order': fields_order
                }
                send_message(user_id, f"✅ Тип: {chosen_type}\nВведите {fields_order[0]}:")
            else:
                send_message(user_id, f"❌ Неизвестный тип. Доступны: {', '.join(SCHEMA_DOC.keys())}")
            continue
            
        
        # Обработка пошагового создания
        if user_id in user_states and user_states[user_id]['action'] == 'new_doc':
            state = user_states[user_id]
            step = state['step']
            field_name = state['fields_order'][step]
            state['fields'][field_name] = event.text.strip()
            
            if step + 1 < len(state['fields_order']):
                state['step'] += 1
                send_message(user_id, f"Введите {state['fields_order'][state['step']]}:")
            else:
                doc_id = save_document(user_id, state['fields'])
                send_message(user_id, f"✅ Документ #{doc_id} сохранён!\n"
                                      f"Данные: {state['fields']}\n"
                                      f"Команды:\n/Показать {doc_id} — получить текст\n/Скачать {doc_id} — скачать DOCX")
                del user_states[user_id]
            continue

        # Команда: получить документ по ID (текст)
        if text.startswith("/показать "):
            try:
                doc_id = int(text.split()[1])
                doc = get_document(doc_id)
                if doc:
                    # Убираем created_at из вывода, если не нужно
                    display_doc = {k: v for k, v in doc.items() if k != 'created_at'}
                    send_message(user_id, f"📄 Документ #{doc_id}:\n{display_doc}")
                else:
                    send_message(user_id, "❌ Документ не найден")
            except:
                send_message(user_id, "❌ Используйте: /Показать <id>")
            continue

#========================================
        if text.startswith("/скачать "):
            try:
                doc_id = int(text.split()[1])
                doc_data = get_document(doc_id)
                if not doc_data:
                    send_message(user_id, "❌ Документ не найден")
                    continue
                try:
                    file_path = generate_docx(doc_id, doc_data, DocName=doc_data["тип_документа"])
                    send_document(user_id, file_path, title=f"{DocType_Rus[doc_data['тип_документа']]}")
                except FileNotFoundError as e:
                    send_message(user_id, f"❌ Ошибка: шаблон не найден. Поместите файл шаблона в {MAIN_PATH}/{doc_data['тип_документа']}")
                except Exception as e:
                    send_message(user_id, f"❌ Ошибка генерации DOCX: {e}")
            except:
                send_message(user_id, "❌ Используйте: /Скачать <id>")
            continue
#=======================================================
        if text.startswith("/редактировать"):
            try:
                doc_id = int(text.split()[1])
                doc = get_document(doc_id)
                if not doc:
                    send_message(user_id, "❌ Документ не найден")
                    continue

                user_states[user_id] = {
                    'action': 'edit_doc',
                    'doc_id': doc_id,
                    'old_doc': doc,
                    'fields': {k: v for k, v in doc.items() if k != 'created_at'},
                    'step': 0,
                    'fields_order': [k for k in doc.keys() if k != 'created_at']
                }
                first_field = user_states[user_id]['fields_order'][0]
                send_message(user_id, f"✏️ Редактирование #{doc_id}\n"
                                      f"Текущие данные: {user_states[user_id]['fields']}\n"
                                      f"Введите новое значение для '{first_field}' (или 'пропустить'):")
            except:
                send_message(user_id, "❌ Используйте: /Редактировать <id>")
            continue

# Обработка редактирования (по шагам)
        if user_id in user_states and user_states[user_id]['action'] == 'edit_doc':
            state = user_states[user_id]
            step = state['step']
            field_name = state['fields_order'][step]

            if event.text.strip().lower() != "пропустить":
                state['fields'][field_name] = event.text.strip()

            if step + 1 < len(state['fields_order']):
                state['step'] += 1
                next_field = state['fields_order'][state['step']]
                send_message(user_id, f"Введите новое значение для '{next_field}' (или 'пропустить'):")
            else:
                if update_document(state['doc_id'], state['fields']):
                    send_message(user_id, f"✅ Документ #{state['doc_id']} обновлён!\n"
                                          f"Новые данные: {state['fields']}")
                else:
                    send_message(user_id, "❌ Ошибка обновления")
                del user_states[user_id]
            continue

        #Команда - список документов
        if text == "/список":
            docs = get_user_docs(user_id)
            if docs:
                msg = "📂 Ваши документы:\n" + "\n".join([f"#{did}: {doc}" for did, doc in docs])
            else:
                msg = "У вас пока нет документов. Напишите 'начать' чтобы создать."
            send_message(user_id, msg)
            continue

#===========================================================================
        if text in ("/help", "помощь"):
            help_text = (
                "📌 Команды:\n"
                "Начать / новый — создать документ\n"
                "/Показать <id> — показать текст документа\n"
                "/Скачать <id> — скачать документ в формате DOCX\n"
                "/Редактировать <id> — редактировать документ\n"
                "/Список — список ваших документов\n"
                "помощь — это сообщение"
            )
            send_message(user_id, help_text)
            continue

        send_message(user_id, "Неизвестная команда. Напишите 'помощь'.")