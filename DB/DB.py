"""
create_wagon_db.py
Создаёт базу данных «Система управления ремонта вагонного оборудования».

Таблицы
--------
вагоны           – карточки вагонов (№, собственник, подразделение, дата ремонта)
договоры         – договоры (№, дата)
услуги           – перечень услуг с ценами
договорные_услуги – связь договоров и услуг
исполнители      – исполнители (ФИО)
выполненные_работы – фактически выполненные работы с привязкой к вагону, договору,
                   услуге и исполнителю, а также интервалом дат и подписантом
"""

import sqlite3
from pathlib import Path


def create_db(db_path: str | Path = "wagons.db") -> None:
    """Создаёт (или пере-создаёт) файл БД со всеми нужными таблицами."""
    db_path = Path(db_path)

    # подключаемся и включаем поддержку внешних ключей
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA foreign_keys = ON;")

    schema = """
    /* ===== Справочники ===== */
    CREATE TABLE IF NOT EXISTS вагоны (
        id             INTEGER PRIMARY KEY,
        номер          TEXT    UNIQUE NOT NULL,  -- 024-06064 и т. д.
        собственник    TEXT,                     -- ДОСС, ФПК …
        подразделение  TEXT,                     -- ЛВЧ, ЛВЧД …
        дата_кр        DATE,                     -- дата капитального ремонта
        дата_кр1       DATE,                     -- новая дата капитального ремонта 1
        дата_квр       DATE,                     -- дата капитально-восстановительного ремонта
        дата_др        DATE                      -- дата деповского ремонта
    );

    CREATE TABLE IF NOT EXISTS договоры (
        id            INTEGER PRIMARY KEY,
        номер         TEXT UNIQUE NOT NULL,      -- № 2024.288648 …
        дата DATE
    );

    CREATE TABLE IF NOT EXISTS исполнители (
        id        INTEGER PRIMARY KEY,
        фио       TEXT UNIQUE NOT NULL           -- ФИО исполнителя
    );

    /* ===== Услуги и цены ===== */
    CREATE TABLE IF NOT EXISTS услуги (
        id             INTEGER PRIMARY KEY,
        наименование   TEXT    UNIQUE NOT NULL,  -- Наименование услуги
        стоимость_без_ндс REAL NOT NULL,         -- базовая стоимость без НДС
        стоимость_с_ндс  REAL NOT NULL,          -- базовая стоимость с НДС
        стоимость_работнику REAL NOT NULL,       -- стоимость работнику за выполнение услуги
        описание       TEXT                       -- подробное описание услуги
    );

    /* ===== Связь договоров и услуг ===== */
    CREATE TABLE IF NOT EXISTS договорные_услуги (
        id_договора   INTEGER NOT NULL,
        id_услуги     INTEGER NOT NULL,
        FOREIGN KEY (id_договора) REFERENCES договоры(id)
            ON UPDATE CASCADE ON DELETE CASCADE,
        FOREIGN KEY (id_услуги) REFERENCES услуги(id)
            ON UPDATE CASCADE ON DELETE CASCADE,
        PRIMARY KEY (id_договора, id_услуги)
    );

    /* ===== Выполненные работы ===== */
    CREATE TABLE IF NOT EXISTS выполненные_работы (
        id          INTEGER PRIMARY KEY,
        id_вагона   INTEGER NOT NULL,
        id_договора INTEGER NOT NULL,
        id_услуги   INTEGER NOT NULL,
        id_исполнителя INTEGER NOT NULL,
        дата_начала_ DATETIME,                  -- начало интервала
        дата_окончания_ DATETIME,               -- конец интервала
        подписант   TEXT,                      -- кто подписывает
        FOREIGN KEY (id_вагона) REFERENCES вагоны(id),
        FOREIGN KEY (id_договора) REFERENCES договоры(id),
        FOREIGN KEY (id_услуги) REFERENCES услуги(id),
        FOREIGN KEY (id_исполнителя) REFERENCES исполнители(id)
    );
    """

    conn.executescript(schema)
    conn.commit()
    conn.close()
    print(f"База данных создана: {db_path.resolve()}")


if __name__ == "__main__":
    create_db()
