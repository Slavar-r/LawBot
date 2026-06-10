import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from vk_api.upload import VkUpload
import sqlite3
import json
import os
from datetime import datetime
from docxtpl import DocxTemplate

# -------------- НАСТРОЙКИ -----------------
VK_TOKEN = rf"{os.getenv('VK_Token')}"
MAIN_PATH= rf"{os.getenv('Bot_Path')}/Doc_sample"

SAMPLES=dict({
              "об увольнении":"firing","о трудоустройстве":"getWork","об отпуске":"vacancy","Извещение об отпуске":"vacancyNotify","Регламент отпуска":"vacancyNotifyQuest",
              "Увольнение":"firing","трудоустройство":"getWork","отпуск":"vacancy","Оповещение об отпуске":"vacancyNotify","Вопрос оповещения об отпуске":"vacancyNotifyQuest"
            })
DocType_Rus=dict({
                  "firing":"заявление на увольнение","об увольнении":"заявление на увольнение",
                  "getWork":"заявление на трудоустройство","о трудоустройстве":"заявление на трудоустройство",
                  "vacancy":"заявление на отпуск","об отпуске":"заявление на отпуск",
                  "vN":"оповещение об отпуске","vacancyNotify":"оповещение об отпуске","извещение об отпуске":"оповещение об отпуске",
                  "vNC":"уведомление об отпуске", "vacancyNotifyQuest":"Уведомление об отпуске","регламент отпуска":"уведомление об отпуске"
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
    "об увольнении":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["ДолжностьОтправителя"],DOCS_VARIABLES["Отправитель"],DOCS_VARIABLES["ДатаРасторжения"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["Кадровик"]],
    "о трудоустройстве":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["ДолжностьОтправителя"],DOCS_VARIABLES["Отправитель"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["Цена"],DOCS_VARIABLES["ЖелаемаяДолжность"],DOCS_VARIABLES["ИспытательныйСрок"],DOCS_VARIABLES["ДатаПриемаНаРаботу"]],
    "об отпуске":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["ДолжностьОтправителя"],DOCS_VARIABLES["Отправитель"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["Причина"],DOCS_VARIABLES["ПланНачОтдых"],DOCS_VARIABLES["ПланКонОтдых"],DOCS_VARIABLES["НачОтдых"],DOCS_VARIABLES["КонОтдых"],DOCS_VARIABLES["КолВоДней"]],
    "извещение об отпуске":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["НачОтдых"],DOCS_VARIABLES["КонОтдых"],DOCS_VARIABLES["Кадровик"], DOCS_VARIABLES["КолВоДней"]],
    "регламент отпуска":[DOCS_VARIABLES["ДолжностьПолучателя"],DOCS_VARIABLES["Получатель"],DOCS_VARIABLES["Организация"],DOCS_VARIABLES["Отдел"],DOCS_VARIABLES["Дата"],DOCS_VARIABLES["ОзнакУвед"],DOCS_VARIABLES["Кадровик"]]
              }
class Policy ():
    def __init__(self):
        self.policy = False
        return
    def ChangeValue(self, value: bool):
        if self.policy != value:
            self.policy = value
            return ("Политика успешно изменена")
        elif self.policy == value:
            return ("Политика не была обновлена: Значение уже установлено")
    def GetValue(self):
        return self.policy
    def Request(self):
        if self.policy == True:
            return "принято"
        else:
            return "не принято"

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
        case "об увольнении":
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
        case "о трудоустройстве":
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
        case "об отпуске":
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
        case "извещение об отпуске":
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
        case "регламент отпуска":
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
def policy_block(text):
    if text.strip().lower() in ("статус", "status"):
        send_message(user_id,policy.Request())
    elif text.strip().lower() in ("принять", "да"):
        send_message(user_id,policy.ChangeValue(True))
    elif text.strip().lower() in ("отозвать", "нет"):
        send_message(user_id,policy.ChangeValue(False))

init_db()
print("Бот запущен")

user_states = {}
policy = Policy()
for event in longpoll.listen():
    if event.type == VkEventType.MESSAGE_NEW and event.to_me:
        user_id = event.user_id
        text = event.text.strip().lower()

        # Команда: начать создание документа
        if text in ("начать", "новый"):
            if policy.GetValue() == False:
                send_message(user_id,"Данный бот работает с личными данными. Чтобы продолжить пользоваться примите условия соглашения")
            user_states[user_id] = {'action': 'policy'}
            types_list = "\n".join([f"- {t.capitalize()}" for t in SCHEMA_DOC.keys()])
            send_message(user_id, f"📌 Выберите тему документа:\n{types_list} \n Или \n Пользовательское соглашение")
            continue


        if user_id in user_states and user_states[user_id].get('action') == 'policy':
            if text.strip().lower() == "пользовательское соглашение":
                send_message(user_id, "Раздел пользовательское соглашение")
                send_message(user_id, f"Принять/Отозвать/Статус")
                policy_block(text)
            user_states[user_id] = {'action': 'choose_type'}


        # --- Обработка выбора типа ---
        if user_id in user_states and user_states[user_id].get('action') == 'choose_type':
            chosen_type = event.text.strip().lower()
            if text.strip().lower() == "пользовательское соглашение":
                user_states[user_id] = {'action': 'policy'}
            elif text.strip().lower() in ("статус", "status"):
                send_message(user_id,policy.Request())
            elif text.strip().lower() in ("принять", "да"):
                send_message(user_id,policy.ChangeValue(True))
            elif text.strip().lower() in ("отозвать", "нет"):
                send_message(user_id,policy.ChangeValue(False))
            elif chosen_type in SCHEMA_DOC and policy.GetValue() == True:
                fields_order = SCHEMA_DOC[chosen_type]
                user_states[user_id] = {
                    'action': 'new_doc',
                    'fields': {'тип_документа': chosen_type},
                    'step': 0,
                    'fields_order': fields_order
                }
                send_message(user_id, f"✅ Тип: {chosen_type}\nВведите {fields_order[0]}:")
            elif chosen_type in SCHEMA_DOC and policy.GetValue() == False:
                send_message(user_id, "❌Пользовательское соглашение не принято")
            elif chosen_type == "прервать":
                del user_states[user_id]
                send_message(user_id, "Успешно прервано")
                continue
            else:
                send_message(user_id, f"❌ Неизвестный тип. Доступны: {', '.join([str(i).capitalize() for i in SCHEMA_DOC.keys()])}, Пользовательское соглашение")
            continue
            
        
        # Обработка пошагового создания
        if user_id in user_states and user_states[user_id]['action'] == 'new_doc' :
            state = user_states[user_id]
            step = state['step']
            field_name = state['fields_order'][step]
            if text.strip().lower() != "прервать":
                state['fields'][field_name] = event.text.strip()
            else:
                del user_states[user_id]
                send_message(user_id, "Успешно прервано")
                continue

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
        if text.startswith("/показать ") and policy.GetValue() == True:
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
        elif text.startswith("/показать ") and policy.GetValue() == False:
            send_message(user_id, f"❌Пользовательское соглашение не принято")
        



#========================================
        if text.startswith("/скачать ") and policy.GetValue() == True:
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
        elif text.startswith("/скачать ") and policy.GetValue() == False:
            send_message(user_id, f"❌Пользовательское соглашение не принято")
#=======================================================
        if text.startswith("/редактировать") and policy.GetValue() == True:
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
                                      f"Введите новое значение для '{first_field}' (также 'пропустить' или 'прервать'):")
            except:
                send_message(user_id, "❌ Используйте: /Редактировать <id>")
            continue
        elif text.startswith("/редактировать ") and policy.GetValue() == False:
            send_message(user_id, f"❌Пользовательское соглашение не принято")

# Обработка редактирования (по шагам)
        if user_id in user_states and user_states[user_id]['action'] == 'edit_doc':
            state = user_states[user_id]
            step = state['step']
            field_name = state['fields_order'][step]

            if event.text.strip().lower() == "пропустить":
                pass
            elif event.text.strip().lower() == "прервать":
                del user_states[user_id]
                send_message(user_id, "Успешно прервано")
                continue
            else:
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
        if text == "/список" and policy.GetValue() == True:
            docs = get_user_docs(user_id)
            if docs:
                msg = "📂 Ваши документы:\n" + "\n".join([f"#{did}: {doc}" for did, doc in docs])
            else:
                msg = "У вас пока нет документов. Напишите 'начать' чтобы создать."
            send_message(user_id, msg)
            continue
        elif text == "/список"and policy.GetValue() == False:
            send_message(user_id, f"❌Пользовательское соглашение не принято")
#===========================================================================
        if text in ("/help", "помощь"):
            help_text = (
                "📌 Команды:\n"
                "Начать / новый — создать документ, пользовательское соглашение\n"
                "/Показать <id> — показать текст документа\n"
                "/Скачать <id> — скачать документ в формате DOCX\n"
                "/Редактировать <id> — редактировать документ\n"
                "/Список — список ваших документов\n"
                "помощь — это сообщение"
            )
            send_message(user_id, help_text)
            continue

        send_message(user_id, "Неизвестная команда. Напишите 'помощь'.")