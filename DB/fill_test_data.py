import sqlite3
from datetime import datetime, timedelta
import random

def fill_test_data(db_path="wagons.db"):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # Очищаем таблицы
    cursor.execute("DELETE FROM выполненные_работы")
    cursor.execute("DELETE FROM договорные_услуги")
    cursor.execute("DELETE FROM услуги")
    cursor.execute("DELETE FROM договоры")
    cursor.execute("DELETE FROM вагоны")
    cursor.execute("DELETE FROM исполнители")

    # Списки тестовых данных
    wagon_numbers = [
        "024-06064", "024-06065", "024-06066", "024-06067", "024-06068",
        "024-06069", "024-06070", "024-06071", "024-06072", "024-06073"
    ]
    
    owners = ["ДОСС", "ФПК", "РЖД", "ТрансКонтейнер"]
    divisions = ["ЛВЧ-1", "ЛВЧ-2", "ЛВЧ-3", "ЛВЧ-4", "ЛВЧ-5"]
    
    workers = [
        "Иванов Иван Иванович",
        "Петров Петр Петрович",
        "Сидоров Сидор Сидорович",
        "Смирнов Алексей Владимирович",
        "Козлов Дмитрий Сергеевич",
        "Николаев Николай Николаевич",
        "Васильев Василий Васильевич",
        "Алексеев Алексей Алексеевич"
    ]
    
    services = [
        {
            "наименование": "Ремонт тормозной системы",
            "стоимость_без_ндс": 15000,
            "стоимость_работнику": 5000,
            "описание": "Полная диагностика и ремонт тормозной системы вагона"
        },
        {
            "наименование": "Замена колесных пар",
            "стоимость_без_ндс": 25000,
            "стоимость_работнику": 8000,
            "описание": "Замена изношенных колесных пар на новые"
        },
        {
            "наименование": "Проверка автосцепки",
            "стоимость_без_ндс": 8000,
            "стоимость_работнику": 3000,
            "описание": "Проверка и регулировка автосцепного устройства"
        },
        {
            "наименование": "Ремонт системы отопления",
            "стоимость_без_ндс": 12000,
            "стоимость_работнику": 4000,
            "описание": "Ремонт и наладка системы отопления пассажирского вагона"
        },
        {
            "наименование": "Проверка электрооборудования",
            "стоимость_без_ндс": 10000,
            "стоимость_работнику": 3500,
            "описание": "Комплексная проверка электрооборудования вагона"
        },
        {
            "наименование": "Ремонт дверей",
            "стоимость_без_ндс": 7000,
            "стоимость_работнику": 2500,
            "описание": "Ремонт и регулировка дверных механизмов"
        },
        {
            "наименование": "Проверка системы вентиляции",
            "стоимость_без_ндс": 9000,
            "стоимость_работнику": 3000,
            "описание": "Проверка и очистка системы вентиляции"
        },
        {
            "наименование": "Ремонт системы водоснабжения",
            "стоимость_без_ндс": 11000,
            "стоимость_работнику": 3800,
            "описание": "Ремонт и прочистка системы водоснабжения"
        }
    ]

    # Заполняем вагоны
    for number in wagon_numbers:
        # Генерируем случайные даты ремонта, некоторые оставляем пустыми
        dates = []
        for _ in range(4):  # Было 3, теперь 4
            if random.random() < 0.7:  # 70% шанс что дата будет заполнена
                dates.append((datetime.now() - timedelta(days=random.randint(0, 365))).strftime("%Y-%m-%d"))
            else:
                dates.append(None)
        
        cursor.execute(
            "INSERT INTO вагоны (номер, собственник, подразделение, дата_кр, дата_кр1, дата_квр, дата_др) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (number, random.choice(owners), random.choice(divisions), dates[0], dates[1], dates[2], dates[3])
        )

    # Заполняем исполнителей
    for worker in workers:
        cursor.execute(
            "INSERT INTO исполнители (фио) VALUES (?)",
            (worker,)
        )

    # Заполняем услуги
    for service in services:
        cost_with_vat = int(service["стоимость_без_ндс"] * 1.2)
        cursor.execute(
            "INSERT INTO услуги (наименование, стоимость_без_ндс, стоимость_с_ндс, стоимость_работнику, описание) VALUES (?, ?, ?, ?, ?)",
            (service["наименование"], service["стоимость_без_ндс"], cost_with_vat, service["стоимость_работнику"], service["описание"])
        )

    # Заполняем договоры
    for i in range(1, 6):
        cursor.execute(
            "INSERT INTO договоры (номер, дата) VALUES (?, ?)",
            (f"2024.{i:06d}", (datetime.now() - timedelta(days=random.randint(0, 180))).strftime("%Y-%m-%d"))
        )

    # Заполняем договорные_услуги
    cursor.execute("SELECT id FROM договоры")
    contract_ids = [row[0] for row in cursor.fetchall()]
    
    cursor.execute("SELECT id FROM услуги")
    service_ids = [row[0] for row in cursor.fetchall()]
    
    for contract_id in contract_ids:
        # Для каждого договора выбираем случайное количество услуг
        selected_services = random.sample(service_ids, random.randint(3, len(service_ids)))
        for service_id in selected_services:
            cursor.execute(
                "INSERT INTO договорные_услуги (id_договора, id_услуги) VALUES (?, ?)",
                (contract_id, service_id)
            )

    # Заполняем выполненные работы
    cursor.execute("SELECT id FROM вагоны")
    wagon_ids = [row[0] for row in cursor.fetchall()]
    
    cursor.execute("SELECT id FROM исполнители")
    worker_ids = [row[0] for row in cursor.fetchall()]
    
    signers = ["Главный инженер", "Начальник депо", "Технический директор", "Руководитель участка"]

    for _ in range(50):  # Создаем 50 выполненных работ
        wagon_id = random.choice(wagon_ids)
        contract_id = random.choice(contract_ids)
        service_id = random.choice(service_ids)
        worker_id = random.choice(worker_ids)
        
        date = datetime.now() - timedelta(days=random.randint(0, 30))
        start_time = datetime.combine(date.date(), datetime.min.time()) + timedelta(hours=random.randint(8, 12))
        end_time = start_time + timedelta(hours=random.randint(1, 8))
        
        cursor.execute(
            """INSERT INTO выполненные_работы 
            (id_вагона, id_договора, id_услуги, id_исполнителя, 
             дата_начала_, дата_окончания_, подписант)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (wagon_id, contract_id, service_id, worker_id,
             start_time.strftime("%Y-%m-%d %H:%M"),
             end_time.strftime("%Y-%m-%d %H:%M"),
             random.choice(signers))
        )

    conn.commit()
    conn.close()
    print("База данных успешно заполнена тестовыми данными")

if __name__ == "__main__":
    fill_test_data() 